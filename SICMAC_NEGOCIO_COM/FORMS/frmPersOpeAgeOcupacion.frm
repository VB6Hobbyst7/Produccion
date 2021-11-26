VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersOpeAgeOcupacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones de Persona Agencia Ocupacion"
   ClientHeight    =   5805
   ClientLeft      =   2985
   ClientTop       =   2895
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8760
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   6855
      Begin SICMACT.FlexEdit grdSectorCIIU 
         Height          =   4605
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   8123
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-OK-Serctor-Monto-nDesde-nHasta"
         EncabezadosAnchos=   "300-320-4100-1700-0-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-R-L-L"
         FormatosEdit    =   "0-0-0-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   7080
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cdmSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   4440
         Width           =   1215
      End
   End
   Begin MSDataListLib.DataCombo dcAgencia 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   210
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Agencia"
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
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPersOpeAgeOcupacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cdmSalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Me.cmdGrabar.Enabled = False
    Me.grdSectorCIIU.Enabled = False
     Me.cmdEditar.Enabled = True
End Sub

Private Sub cmdEditar_Click()
    Me.cmdGrabar.Enabled = True
    Me.grdSectorCIIU.Enabled = True
    Me.cmdEditar.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
   Dim i As Integer
   Dim items As Integer
   Dim objCOMDPersona As COMDPersona.DCOMPersonas
   items = 0
   For i = 1 To Me.grdSectorCIIU.Rows - 1
        If Me.grdSectorCIIU.TextMatrix(i, 1) <> "" Then
            items = 1
        End If
   Next i
   If items = 0 Then
        MsgBox "Debe Ingresar Datos en al menos un Sector"
        Exit Sub
   End If
    
    Dim sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    Set objCOMDPersona = New COMDPersona.DCOMPersonas
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   
   If MsgBox("Se van ha guardar los Datos", vbYesNo, "AVISO") = vbYes Then
        objCOMDPersona.modificarParamPersAgeOcupacion dcAgencia.BoundText
        For i = 1 To Me.grdSectorCIIU.Rows - 1
                 objCOMDPersona.insertarParamPersAgeOcupacion dcAgencia.BoundText, CInt(grdSectorCIIU.TextMatrix(i, 4)), CInt(grdSectorCIIU.TextMatrix(i, 5)), CCur(grdSectorCIIU.TextMatrix(i, 3)), sMovNro, CInt(IIf(grdSectorCIIU.TextMatrix(i, 1) = "", 0, 1))
                  
        Next i
        MsgBox "Se Guardaron Todos los Datos"
    End If
   Me.cmdGrabar.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.grdSectorCIIU.Enabled = False
End Sub

Private Sub Form_Load()
    CargarAgencias
    CargarSectoresCIIU
End Sub
Private Sub CargarAgencias()
    Dim rsAgencia As New ADODB.Recordset
    Dim objCOMNCredito As COMNCredito.NCOMBPPR
    
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsAgencia.DataSource = objCOMNCredito.getCargarAgencias
    dcAgencia.BoundColumn = "cAgeCod"
    dcAgencia.DataField = "cAgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = 0
End Sub
Private Sub CargarSectoresCIIU()
    Dim rsCIIU As New ADODB.Recordset
    Dim objCOMDPersona As COMDPersona.DCOMPersonas
    
    Set objCOMDPersona = New COMDPersona.DCOMPersonas
    Set rsCIIU.DataSource = objCOMDPersona.CargarSectoresCIUU
    
    
    If Not (rsCIIU.EOF And rsCIIU.BOF) Then
        grdSectorCIIU.Rows = rsCIIU.RecordCount + 1
        Dim i As Integer
        
        For i = 1 To rsCIIU.RecordCount
            grdSectorCIIU.TextMatrix(i, 0) = i
            grdSectorCIIU.TextMatrix(i, 1) = 1
            grdSectorCIIU.TextMatrix(i, 2) = rsCIIU!cDescrip
            grdSectorCIIU.TextMatrix(i, 3) = 0
            grdSectorCIIU.TextMatrix(i, 4) = rsCIIU!nDesde
            grdSectorCIIU.TextMatrix(i, 5) = rsCIIU!nHasta
            rsCIIU.MoveNext
        Next i
    End If
End Sub
Private Sub dcAgencia_Change()
    Me.grdSectorCIIU.Enabled = False
    If dcAgencia.BoundText <> "0" Then
        
        cargarParametrosAgencia
       
        Me.cmdEditar.Enabled = True
        Me.cmdCancelar.Enabled = True
    Else
        
        Me.cmdEditar.Enabled = False
        Me.cmdCancelar.Enabled = False
        Me.cmdGrabar.Enabled = False
    End If
    
End Sub
Private Sub cargarParametrosAgencia()
    Dim rsCIIU As New ADODB.Recordset
    Dim objCOMDPersona As COMDPersona.DCOMPersonas
    Dim i As Integer
    
    Set objCOMDPersona = New COMDPersona.DCOMPersonas
    Set rsCIIU.DataSource = objCOMDPersona.CargarParamAgeOcupacion(dcAgencia.BoundText)
    
    
    If Not (rsCIIU.EOF And rsCIIU.BOF) Then
        
        For i = 1 To rsCIIU.RecordCount
             
            grdSectorCIIU.TextMatrix(i, 1) = rsCIIU!nActivo
            grdSectorCIIU.TextMatrix(i, 3) = Format(rsCIIU!nMonto, "##,##0.00")
            
            rsCIIU.MoveNext
        Next i
    Else
        For i = 1 To grdSectorCIIU.Rows - 1
            grdSectorCIIU.TextMatrix(i, 1) = 0
            grdSectorCIIU.TextMatrix(i, 3) = Format("0", "##,##0.00")
        Next i
    End If
End Sub

Private Sub grdSectorCIIU_Click()
    With grdSectorCIIU
        If .Col = 1 Then
            .TextMatrix(.Row, 3) = Format("0", "##,##0.00")
             
        End If
    End With
End Sub


Private Sub grdSectorCIIU_OnCellChange(pnRow As Long, pnCol As Long)
    With grdSectorCIIU
        If .TextMatrix(.Row, 1) = "" Then
          .TextMatrix(.Row, .Col) = Format("0", "##,##0.00")
        End If
    End With
End Sub


