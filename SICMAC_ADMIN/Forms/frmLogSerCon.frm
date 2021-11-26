VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogSerCon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio/Contratación : Registro"
   ClientHeight    =   5910
   ClientLeft      =   720
   ClientTop       =   1560
   ClientWidth     =   9180
   Icon            =   "frmLogSerCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGarantia 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   5040
      TabIndex        =   25
      Top             =   180
      Visible         =   0   'False
      Width           =   3930
      Begin VB.CommandButton cmdSerTpoGar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1545
         TabIndex        =   32
         Top             =   4395
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtGarNro 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   27
         Top             =   2415
         Width           =   2235
      End
      Begin VB.TextBox txtGarComenta 
         Height          =   1380
         Left            =   120
         MaxLength       =   70
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2985
         Width           =   3705
      End
      Begin Sicmact.FlexEdit fgeGarantia 
         Height          =   2010
         Left            =   150
         TabIndex        =   28
         Top             =   345
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   3545
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Opc"
         EncabezadosAnchos=   "400-0-2500-400"
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
         ColumnasAEditar =   "X-X-X-3"
         ListaControles  =   "0-0-0-5"
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
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
         Caption         =   "Comentario"
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
         Index           =   8
         Left            =   165
         TabIndex        =   31
         Top             =   2775
         Width           =   990
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Documento"
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
         Index           =   6
         Left            =   165
         TabIndex        =   30
         Top             =   2445
         Width           =   1050
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo de garantía"
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
         Index           =   7
         Left            =   150
         TabIndex        =   29
         Top             =   105
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdDistribuye 
      Caption         =   ">>>"
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
      Height          =   285
      Left            =   6285
      MaskColor       =   &H8000000F&
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton cmdSerTpoDis 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   8010
      TabIndex        =   21
      Top             =   4500
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdSerTpoDis 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   7065
      TabIndex        =   20
      Top             =   4500
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdSerTpoDis 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   6120
      TabIndex        =   19
      Top             =   4500
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame fraRegistro 
      Enabled         =   0   'False
      Height          =   1845
      Left            =   225
      TabIndex        =   11
      Top             =   3990
      Visible         =   0   'False
      Width           =   5805
      Begin Sicmact.TxtBuscar txtNumero 
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Top             =   345
         Width           =   1740
         _ExtentX        =   3069
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1860
         TabIndex        =   3
         Top             =   345
         Width           =   3780
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   135
         MaxLength       =   70
         TabIndex        =   4
         Top             =   885
         Width           =   5490
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3810
         MaxLength       =   12
         TabIndex        =   7
         Top             =   1440
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   1425
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   37116
         MaxDate         =   73415
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   315
         Left            =   1935
         TabIndex        =   6
         Top             =   1425
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   37116
         MaxDate         =   73415
         MinDate         =   36526
      End
      Begin VB.Label lblEtiqueta 
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
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   150
         Width           =   765
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Descripción"
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
         Left            =   180
         TabIndex        =   15
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha Inicio"
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
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   1215
         Width           =   1200
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha Final"
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
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   13
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Monto"
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
         Height          =   180
         Index           =   5
         Left            =   3840
         TabIndex        =   12
         Top             =   1230
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdSerTpo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   6165
      TabIndex        =   8
      Top             =   4500
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdSerTpo 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   6165
      TabIndex        =   9
      Top             =   4845
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdSerTpo 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   6165
      TabIndex        =   10
      Top             =   5190
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7605
      TabIndex        =   22
      Top             =   5340
      Width           =   1305
   End
   Begin VB.ComboBox cboSerTpo 
      Height          =   315
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2100
   End
   Begin Sicmact.FlexEdit fgeSerTpo 
      Height          =   3480
      Left            =   225
      TabIndex        =   2
      Top             =   525
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6138
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-nMovNro-Código-Nombre-Descripción-Fecha Inicio-Fecha Final-Monto"
      EncabezadosAnchos=   "400-0-0-1500-3400-1000-1000-1000"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "L-L-L-L-L-C-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-2"
      TextArray0      =   "Item"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   0
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin Sicmact.FlexEdit fgeDistribucion 
      Height          =   3915
      Left            =   6105
      TabIndex        =   17
      Top             =   525
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6906
      Cols0           =   3
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Fecha-Monto"
      EncabezadosAnchos=   "400-1000-1000"
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
      ColumnasAEditar =   "X-1-2"
      ListaControles  =   "0-2-0"
      EncabezadosAlineacion=   "L-C-R"
      FormatosEdit    =   "0-0-2"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   0
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label lblDistribucion 
      Caption         =   "Distribución"
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
      Left            =   6225
      TabIndex        =   18
      Top             =   270
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Tipo de Servicio :"
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
      Left            =   300
      TabIndex        =   1
      Top             =   165
      Width           =   1530
   End
End
Attribute VB_Name = "frmLogSerCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pnFrmTpo As Integer

Public Sub Inicio(ByVal pnFormTipo As Integer)
'1 - Registro ;  2 - Distribución ; 3 -
pnFrmTpo = pnFormTipo
Me.Show 1
End Sub

Private Sub cboSerTpo_Click()
    Dim rs As ADODB.Recordset
    Dim clsDSC  As DLogSerCon
    
    If cboSerTpo.ListCount > 0 Then
        Call Limpiar
        fgeSerTpo.Clear
        fgeSerTpo.FormaCabecera
        fgeSerTpo.Rows = 2
        If pnFrmTpo = 2 Then
            fgeDistribucion.Clear
            fgeDistribucion.FormaCabecera
            fgeDistribucion.Rows = 2
        ElseIf pnFrmTpo = 3 Then
            fgeGarantia.Clear
            fgeGarantia.FormaCabecera
            fgeGarantia.Rows = 2
        End If
        If (Val(Right(cboSerTpo.Text, 3)) Mod 10) = 0 Then
            'Si es cabecera
            If pnFrmTpo = 1 Then
                'Registro
                cmdSerTpo(0).Enabled = False
                cmdSerTpo(1).Enabled = False
                cmdSerTpo(2).Enabled = False
            ElseIf pnFrmTpo = 2 Then
                'Distribución
                cmdSerTpoDis(1).Enabled = False
                cmdSerTpoDis(2).Enabled = False
                cmdSerTpoDis(4).Enabled = False
                cmdDistribuye.Enabled = False
            ElseIf pnFrmTpo = 3 Then
                cmdSerTpoGar.Enabled = False
            End If
        Else
            'Detalle
            If pnFrmTpo = 1 Then
                'Registro
                cmdSerTpo(0).Enabled = True
                cmdSerTpo(1).Enabled = False
                cmdSerTpo(2).Enabled = True
            ElseIf pnFrmTpo = 2 Then
                'Distribución
                cmdSerTpoDis(1).Enabled = False
                cmdSerTpoDis(2).Enabled = False
                cmdSerTpoDis(4).Enabled = False
                cmdDistribuye.Enabled = False
            ElseIf pnFrmTpo = 3 Then
                cmdSerTpoGar.Enabled = False
            End If
            
            Set rs = New ADODB.Recordset
            Set clsDSC = New DLogSerCon
            
            Set rs = clsDSC.CargaSerCon(SCTodosTpo, Val(Right(cboSerTpo.Text, 5)))
            If rs.RecordCount > 0 Then
                Set fgeSerTpo.Recordset = rs
                Call fgeSerTpo_OnRowChange(fgeSerTpo.Row, fgeSerTpo.Col)
            End If
        
            Set rs = Nothing
            Set clsDSC = Nothing
        End If
    End If
End Sub

Private Sub cmdDistribuye_Click()
    Dim nSCRow As Integer, nCont As Integer
    Dim dFecha As Date, dFinal As Date
    
    nSCRow = fgeSerTpo.Row
    dFecha = fgeSerTpo.TextMatrix(nSCRow, 5)
    dFinal = fgeSerTpo.TextMatrix(nSCRow, 6)
    
    fgeDistribucion.Clear
    fgeDistribucion.FormaCabecera
    fgeDistribucion.Rows = 2
    
    Do While DateDiff("d", dFecha, dFinal) > 0
        fgeDistribucion.AdicionaFila
        fgeDistribucion.TextMatrix(fgeDistribucion.Row, 1) = dFecha
        fgeDistribucion.TextMatrix(fgeDistribucion.Row, 2) = fgeSerTpo.TextMatrix(nSCRow, 7)
        dFecha = DateAdd("m", 1, dFecha)
    Loop
    fgeDistribucion.lbEditarFlex = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSerTpo_Click(Index As Integer)
    Dim clsDGnral As DLogGeneral
    Dim clsDMov As DLogMov
    Dim nBSRow As Integer
    Dim sSerTpoNro As String, sActualiza As String
    Dim nSerTpoNro As Long
    Dim nSerTpo As Integer, nResult As Integer
    
    'Botones de comandos del Registro
    If Index = 0 Then
        'Nuevo
        Call Limpiar
        
        fraRegistro.Enabled = True
        txtNumero.SetFocus
        
        cmdSerTpo(0).Enabled = False
        cmdSerTpo(1).Enabled = True
        cmdSerTpo(2).Enabled = False
    ElseIf Index = 1 Then
        'GRABAR
        txtDescripcion.Text = Replace(txtDescripcion.Text, "'", "", , , vbTextCompare)
        
        If Len(Trim(txtNumero.Text)) = 0 Then
            MsgBox "Falta determinar la persona", vbInformation, " Aviso "
            Exit Sub
        End If
        If Len(Trim(txtMonto.Text)) = 0 Then
            MsgBox "Falta ingresar el monto", vbInformation, " Aviso "
            Exit Sub
        End If
        
        Set clsDGnral = New DLogGeneral
        
        nSerTpo = Val(Right(cboSerTpo.Text, 5))
        sSerTpoNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        
        
        Set clsDGnral = Nothing
        Set clsDMov = New DLogMov
        
        'Grabación de MOV -MOVREF
        clsDMov.InsertaMov sSerTpoNro, Trim(Str(gLogOpeConRegistro)), "", 0
        nSerTpoNro = clsDMov.GetnMovNro(sSerTpoNro)
        
        'Actualiza LogSerCon
        clsDMov.InsertaSerCon nSerTpoNro, nSerTpo, txtNumero.Text, txtDescripcion.Text, _
            txtMonto.Text, dtpInicio.Value, dtpFinal.Value, sActualiza
        
        'Ejecuta todos los querys en una transacción
        'nResult = clsDMov.EjecutaBatch
        Set clsDMov = Nothing
        
        If nResult = 0 Then
            fraRegistro.Enabled = False
            fgeSerTpo.AdicionaFila
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 1) = nSerTpoNro
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 2) = Trim(txtNumero.Text)
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 3) = Trim(txtNombre.Text)
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 4) = Trim(txtDescripcion.Text)
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 5) = dtpInicio.Value
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 6) = dtpFinal.Value
            fgeSerTpo.TextMatrix(fgeSerTpo.Row, 7) = Format(Trim(txtMonto.Text), "#0.00")
            fgeSerTpo.SetFocus
            
            cmdSerTpo(0).Enabled = True
            cmdSerTpo(1).Enabled = False
            cmdSerTpo(2).Enabled = True
        Else
            MsgBox "Error al grabar la información", vbInformation, " Aviso "
        End If
    ElseIf Index = 2 Then
        'ELIMINAR
        nBSRow = fgeSerTpo.Row
        nSerTpoNro = Val(fgeSerTpo.TextMatrix(nBSRow, 1))
        
        If fgeSerTpo.TextMatrix(nBSRow, 0) <> "" And nSerTpoNro <> 0 Then
            If MsgBox("¿ Estás seguro de eliminar " & fgeSerTpo.TextMatrix(nBSRow, 4) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDMov = New DLogMov
            
                'Actualiza LogSerCon
                clsDMov.EliminaSerCon nSerTpoNro
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    fgeSerTpo.EliminaFila nBSRow
                    Call Limpiar
                    Call fgeSerTpo_OnRowChange(fgeSerTpo.Row, fgeSerTpo.Col)
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
        Else
            cmdSerTpo(2).Enabled = False
        End If
    End If
End Sub

Private Sub cmdSerTpoDis_Click(Index As Integer)
    'Dim clsDGnral As DLogGeneral
    Dim clsDMov As DLogMov
    Dim nDisRow As Integer, nCont As Integer, nResult As Integer
    Dim nSerTpoNro As Long
    'Botones de comandos del Registro
    
    If Index = 0 Then
        'NINGUNO
    ElseIf Index = 1 Then
        'AGREGAR
        fgeDistribucion.AdicionaFila
        fgeDistribucion.lbEditarFlex = True
    ElseIf Index = 2 Then
        'ELIMINAR
        nDisRow = fgeDistribucion.Row
        If fgeDistribucion.TextMatrix(nDisRow, 0) <> "" Then
            fgeDistribucion.EliminaFila nDisRow
        Else
            cmdSerTpoDis(2).Enabled = False
        End If
    ElseIf Index = 3 Then
        'NINGUNO
        
    ElseIf Index = 4 Then
        'GRABAR
        If fgeDistribucion.TextMatrix(1, 0) <> "" Then
            'Validar
            For nCont = 1 To fgeDistribucion.Rows - 1
                If fgeDistribucion.TextMatrix(nCont, 1) = "" Or fgeDistribucion.TextMatrix(nCont, 1) = "" Then
                    MsgBox "Falta completar información en el item " & nCont, vbInformation, " Aviso"
                    Exit Sub
                End If
            Next
            
            If MsgBox("¿ Estás seguro de Grabar esta información ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDMov = New DLogMov
            
                'Inseerta LogSerConDetalle
                nSerTpoNro = Val(fgeSerTpo.TextMatrix(fgeSerTpo.Row, 1))
                For nCont = 1 To fgeDistribucion.Rows - 1
                    clsDMov.InsertaSerConDet nSerTpoNro, _
                        fgeDistribucion.TextMatrix(nCont, 1), CCur(fgeDistribucion.TextMatrix(nCont, 2))
                Next
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdSerTpoDis(1).Enabled = False
                    cmdSerTpoDis(2).Enabled = False
                    cmdSerTpoDis(4).Enabled = False
                    cmdDistribuye.Enabled = False
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
            
        End If
    End If
End Sub

Private Sub cmdSerTpoGar_Click()
    Dim clsDMov As DLogMov
    Dim nCont As Integer, nTpoGar As Integer, nResult As Integer
    Dim nSerTpoNro As Long
    
    'Validar
    txtGarNro.Text = Trim(Replace(txtGarNro.Text, "'", "", , , vbTextCompare))
    txtGarComenta.Text = Trim(Replace(txtGarComenta.Text, "'", "", , , vbTextCompare))
    nTpoGar = 0
    For nCont = 1 To fgeGarantia.Rows - 1
        If fgeGarantia.TextMatrix(nCont, 3) = "." Then
            nTpoGar = Val(fgeGarantia.TextMatrix(nCont, 1))
            Exit For
        End If
    Next
    If nTpoGar = 0 Then
        MsgBox "Falta determinar el tipo de garantía", vbInformation, " Aviso "
        Exit Sub
    End If
    If txtGarNro.Text = "" Then
        MsgBox "Falta determinar el número de la garantía", vbInformation, " Aviso "
        Exit Sub
    End If
    
    If MsgBox("¿ Estás seguro de Grabar esta información ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        Set clsDMov = New DLogMov
    
        'Inserta LogSerConGarantia
        nSerTpoNro = Val(fgeSerTpo.TextMatrix(fgeSerTpo.Row, 1))
        
        clsDMov.InsertaSerConGar nSerTpoNro, _
            nTpoGar, txtGarNro.Text, txtGarComenta.Text
        'Ejecuta todos los querys en una transacción
        'nResult = clsDMov.EjecutaBatch
        Set clsDMov = Nothing
        
        If nResult = 0 Then
            cmdSerTpoGar.Enabled = False
            fraGarantia.Enabled = False
        Else
            MsgBox "Error al grabar la información", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub fgeDistribucion_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim nCont As Integer
    Dim dFecha As Date
    If pnCol = 1 Then
        If (DateDiff("d", fgeDistribucion.TextMatrix(pnRow, 1), fgeSerTpo.TextMatrix(fgeSerTpo.Row, 5)) > 0 Or DateDiff("d", fgeDistribucion.TextMatrix(pnRow, 1), fgeSerTpo.TextMatrix(fgeSerTpo.Row, 6)) < 0) Then
            Cancel = False
        End If
        dFecha = fgeDistribucion.TextMatrix(pnRow, 1)
        For nCont = 1 To fgeDistribucion.Rows - 1
            If DateDiff("d", dFecha, fgeDistribucion.TextMatrix(nCont, 1)) = 0 And pnRow <> nCont Then
                Cancel = False
            End If
        Next
    End If
End Sub

Private Sub fgeSerTpo_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As ADODB.Recordset
    Dim clsDSC As DLogSerCon
    Dim clsDGnral As DLogGeneral
    
    If fgeSerTpo.TextMatrix(pnRow, 0) <> "" Then
        If pnFrmTpo = 1 Then
            'REGISTRO
            txtNumero.Text = fgeSerTpo.TextMatrix(pnRow, 2)
            txtNombre.Text = fgeSerTpo.TextMatrix(pnRow, 3)
            txtDescripcion.Text = fgeSerTpo.TextMatrix(pnRow, 4)
            dtpInicio.Value = fgeSerTpo.TextMatrix(pnRow, 5)
            dtpFinal.Value = fgeSerTpo.TextMatrix(pnRow, 6)
            txtMonto.Text = fgeSerTpo.TextMatrix(pnRow, 7)
        ElseIf pnFrmTpo = 2 Then
            'DISTRIBUCION
            fgeDistribucion.Clear
            fgeDistribucion.FormaCabecera
            fgeDistribucion.Rows = 2
            
            Set rs = New ADODB.Recordset
            Set clsDSC = New DLogSerCon
            
            Set rs = clsDSC.CargaSerConDet(Val(fgeSerTpo.TextMatrix(pnRow, 1)))
            If rs.RecordCount > 0 Then
                Set fgeDistribucion.Recordset = rs
                cmdSerTpoDis(1).Enabled = False
                cmdSerTpoDis(2).Enabled = False
                cmdSerTpoDis(4).Enabled = False
                cmdDistribuye.Enabled = False
            Else
                cmdSerTpoDis(1).Enabled = True
                cmdSerTpoDis(2).Enabled = True
                cmdSerTpoDis(4).Enabled = True
                cmdDistribuye.Enabled = True
            End If
            
            Set rs = Nothing
            Set clsDSC = Nothing
        ElseIf pnFrmTpo = 3 Then
            'GARANTIA
            fgeGarantia.Clear
            fgeGarantia.FormaCabecera
            fgeGarantia.Rows = 2
            txtGarNro.Text = ""
            txtGarComenta.Text = ""
            
            Set rs = New ADODB.Recordset
            Set clsDSC = New DLogSerCon
            Set rs = clsDSC.CargaSerConGar(Val(fgeSerTpo.TextMatrix(pnRow, 1)))
            If rs.RecordCount > 0 Then
                fgeGarantia.EncabezadosAnchos = "400-0-2500-0"
                fraGarantia.Enabled = False
                cmdSerTpoGar.Enabled = False
                fgeGarantia.AdicionaFila
                fgeGarantia.TextMatrix(1, 1) = rs!nLogSerConGarTpo
                fgeGarantia.TextMatrix(1, 2) = rs!cConsDescripcion
                txtGarNro.Text = rs!cLogSerConGarNro
                txtGarComenta.Text = rs!cLogSerConGarDescripcion
            Else
                'Carga garantias para escoger
                fgeGarantia.EncabezadosAnchos = "400-0-2500-400"
                fraGarantia.Enabled = True
                cmdSerTpoGar.Enabled = True
                Set rs = Nothing
                Set clsDGnral = New DLogGeneral
                Set rs = clsDGnral.CargaConstante(gPersGarantia)
                If rs.RecordCount > 0 Then
                    Set fgeGarantia.Recordset = rs
                End If
                Set clsDGnral = Nothing
            End If
            
            Set rs = Nothing
            Set clsDSC = Nothing
        Else
            MsgBox "Tipo no reconocido", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim clsDGnral As DLogGeneral
    Dim rsTpo As ADODB.Recordset

    Call CentraForm(Me)
    
    Set clsDGnral = New DLogGeneral
    Set rsTpo = New ADODB.Recordset
    Set rsTpo = clsDGnral.CargaConstante(gLogSerConTpo, False)
    Do While Not rsTpo.EOF
        If (rsTpo!nConsValor Mod 10) = 0 Then
            cboSerTpo.AddItem rsTpo!cConsDescripcion & Space(40) & rsTpo!nConsValor
        Else
            cboSerTpo.AddItem Space(5) & rsTpo!cConsDescripcion & Space(40) & rsTpo!nConsValor
        End If
        rsTpo.MoveNext
    Loop
    Set rsTpo = Nothing
    
    If pnFrmTpo = 1 Then
        'Registro
        Me.Caption = "Servicio/Contrato : Registro "
        fraRegistro.Visible = True
        cmdSerTpo(0).Visible = True
        cmdSerTpo(1).Visible = True
        cmdSerTpo(2).Visible = True
        
    ElseIf pnFrmTpo = 2 Then
        'Distribución
        Me.Caption = "Servicio/Contrato : Distribución "
        
        cmdDistribuye.Visible = True
        cmdSerTpoDis(1).Visible = True
        cmdSerTpoDis(2).Visible = True
        cmdSerTpoDis(4).Visible = True
        fgeSerTpo.EncabezadosAnchos = "400-0-0-2000-0-1000-1000-1000"
        fgeSerTpo.Width = fgeSerTpo.Width - 2900
        fgeSerTpo.Height = fgeSerTpo.Height + 1500
        lblDistribucion.Visible = True
        fgeDistribucion.Visible = True
    ElseIf pnFrmTpo = 3 Then
        'Garantías
        Me.Caption = "Servicio/Contrato : Garantías "
        
        cmdSerTpoGar.Visible = True
        fgeSerTpo.EncabezadosAnchos = "400-0-0-2000-2000-0-0-0"
        fgeSerTpo.Width = fgeSerTpo.Width - 3900
        fgeSerTpo.Height = fgeSerTpo.Height + 1500
        fraGarantia.Visible = True
    Else
        MsgBox "Tipo de Formulario no reconocido", vbInformation, " Aviso "
        cboSerTpo.Enabled = False
    End If
    
End Sub

Private Sub Limpiar()
    txtNumero.Text = ""
    txtNombre.Text = ""
    txtDescripcion.Text = ""
    txtMonto.Text = ""
    dtpInicio.Value = gdFecSis
    dtpFinal.Value = gdFecSis

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 8, 2)
End Sub

Private Sub txtNumero_EmiteDatos()
    txtNombre.Text = txtNumero.psDescripcion
End Sub
