VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColHistoCambios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Cambios - Créditos Convenio"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   10815
   Icon            =   "frmColHistoCambios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVer1 
      Caption         =   " Parametros de Búsqueda"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   10575
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias..."
         Height          =   345
         Left            =   6720
         TabIndex        =   18
         Top             =   1080
         Width           =   1140
      End
      Begin VB.ComboBox CmbInstitucionDest 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   3705
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optConvenio 
         Caption         =   "Convenio"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdProAsig 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   9360
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox CmbInstitucion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   3585
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   1080
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHasta 
         Height          =   300
         Left            =   4560
         TabIndex        =   17
         Top             =   1080
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblOrigen 
         Caption         =   "Origen:"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDestino 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDesde 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblHasta 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   2293
      TabCaption(0)   =   "Asignación"
      TabPicture(0)   =   "frmColHistoCambios.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flexAsignacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reasignación"
      TabPicture(1)   =   "frmColHistoCambios.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexReasignacion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Retiro"
      TabPicture(2)   =   "frmColHistoCambios.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexRetiro"
      Tab(2).ControlCount=   1
      Begin SICMACT.FlexEdit FlexRetiro 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8493
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Usuario-Fecha-Hora-Conv. Origen"
         EncabezadosAnchos=   "250-2000-1300-1200-1200-3600"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "-0----0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "0--0-0-0-L"
         FormatosEdit    =   "-0----0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FlexReasignacion 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8493
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Usuario-Fecha-Hora-Conv. Origen-Conv. Destino"
         EncabezadosAnchos=   "250-2000-1300-1200-1200-3600-3600"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "-0----0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "0--0-0-0-L-L"
         FormatosEdit    =   "-0----0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flexAsignacion 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8493
         Cols0           =   6
         EncabezadosNombres=   "#-Crédito-Usuario-Fecha-Hora-Conv. Origen"
         EncabezadosAnchos=   "250-2000-1300-1300-1300-3600"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "-0----0"
         EncabezadosAlineacion=   "0--0-0-0-C"
         FormatosEdit    =   "-0----0"
         TextArray0      =   "#"
         ColWidth0       =   255
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmColHistoCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTpoBusca As Integer
Dim nTpoOpe As Integer
Dim sAgencias As String

Private Sub cmdAgencia_Click()
    sAgencias = ""
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdProAsig_Click()
    Dim i As Integer
    Dim nContAge As Integer
    Dim nContAgencias As Integer
    
    Screen.MousePointer = 11
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            sAgencias = sAgencias & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ","
        End If
    Next i
    
    sAgencias = Mid(sAgencias, 1, Len(sAgencias) - 1)

    If optFecha.value = True Then
        If txtDesde.Text = "__/__/____" Or txtHasta.Text = "__/__/____" Then
            MsgBox "Usted Debe ingresar la Fecha valida", vbInformation, "Alerta"
            Exit Sub
        End If
    End If
    If SSTab1.Tab = 0 Then
      Call CargarDatos(flexAsignacion)
    ElseIf SSTab1.Tab = 1 Then
       Call CargarDatos(FlexReasignacion)
    Else
       Call CargarDatos(FlexRetiro)
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub CargaInstitucion()
Dim oPersonas As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset

    On Error GoTo ERRORCarga
    CmbInstitucion.Clear
    CmbInstitucionDest.Clear
    
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set rs = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
    Set oPersonas = Nothing
    
    CmbInstitucion.AddItem Space(30) & " -- TODOS --" & Space(250) & "00"
    CmbInstitucionDest.AddItem Space(30) & "-- TODOS --" & Space(250) & "00"
    
    Do While Not rs.EOF
        CmbInstitucion.AddItem PstaNombre(rs!cPersNombre) & Space(250) & rs!cPersCod
        CmbInstitucionDest.AddItem PstaNombre(rs!cPersNombre) & Space(250) & rs!cPersCod
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    
        
    If CmbInstitucion.ListCount > 0 Then
        CmbInstitucion.ListIndex = 0
        If SSTab1.Tab = 0 Then
            CmbInstitucion.ListIndex = 1
        End If
    End If
    If CmbInstitucionDest.ListCount > 0 Then
        CmbInstitucionDest.ListIndex = 0
    End If
    Exit Sub
ERRORCarga:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub Command1_Click()
    frmSelectAgencias.Show
End Sub

Private Sub Form_Load()
    Call CargaInstitucion
    Call Limpiar
End Sub

Private Sub optConvenio_Click()
    HablitaControl True
    nTpoOpe = 2
    CmbInstitucion.ListIndex = IIf(SSTab1.Tab = 1, 0, 1)
    txtDesde.Text = "__/__/____"
    txtHasta.Text = "__/__/____"
End Sub

Private Sub optFecha_Click()
    HablitaControl False
    nTpoOpe = 3
End Sub

Private Sub optTodos_Click()
    CmbInstitucion.Enabled = False
    CmbInstitucionDest.Enabled = False
    txtDesde.Enabled = False
    txtHasta.Enabled = False
    txtDesde.Text = "__/__/____"
    txtHasta.Text = "__/__/____"
    nTpoOpe = 1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
     If SSTab1.Tab = 1 Then
        CmbInstitucionDest.Visible = True
        lblDestino.Visible = True
        lblOrigen.Visible = True
        CmbInstitucion.ListIndex = 0
    Else
        CmbInstitucionDest.Visible = False
        lblDestino.Visible = False
        lblOrigen.Visible = False
        CmbInstitucion.ListIndex = 1
    End If
End Sub
Public Sub Limpiar()
    SSTab1.Tab = 0
    CmbInstitucionDest.Visible = False
    lblDestino.Visible = False
    lblOrigen.Visible = False
    nTpoBusca = 1
    nTpoOpe = 1
    sAgencias = ""
End Sub
Public Sub HablitaControl(ByVal pbHabilita As Boolean)
        CmbInstitucion.Enabled = pbHabilita
        CmbInstitucionDest.Enabled = pbHabilita
        txtDesde.Enabled = Not pbHabilita
        txtHasta.Enabled = Not pbHabilita
End Sub

Public Sub CargarDatos(ByVal poControl As FlexEdit)
Dim oCredito As New COMDCredito.DCOMCredito
  Dim oDrDatos As New ADODB.Recordset
  Dim i As Integer
  
  Set oCredito = New COMDCredito.DCOMCredito
    
  
    
    nTpoBusca = SSTab1.Tab + 1
    
    If nTpoBusca = 2 And optConvenio.value = True Then
        Dim nBusq As Integer
        If optConvenio.value = True Then
            If Trim(Right(CmbInstitucion.Text, 20)) = "00" And Trim(Right(CmbInstitucionDest.Text, 20)) = "00" Then
                MsgBox "Debe seleccionar al menos un convenio origen o un convenio destino", vbCritical, "Alerta"
                Exit Sub
            End If
        End If
        
        If Trim(Right(CmbInstitucion.Text, 20)) = "00" Then
            nBusq = 2
        ElseIf Trim(Right(CmbInstitucionDest.Text, 20)) = "00" Then
            nBusq = 1
        Else
            nBusq = 3
        End If
        
        Set oDrDatos = oCredito.HistorialConvenioReasigna(Trim(Right(CmbInstitucion.Text, 20)), Trim(Right(CmbInstitucionDest.Text, 20)), nBusq)
    Else
        Set oDrDatos = oCredito.ObtenerHistorialConvenio(Trim(Right(CmbInstitucion.Text, 20)), Trim(Right(CmbInstitucionDest.Text, 20)), _
        IIf(txtDesde.Text = "__/__/____", "", txtDesde.Text), IIf(txtHasta.Text = "__/__/____", "", txtHasta.Text), nTpoBusca, nTpoOpe, sAgencias)
    End If

        poControl.Clear
        FormateaFlex poControl
    For i = 1 To oDrDatos.RecordCount
        
        poControl.AdicionaFila
        poControl.TextMatrix(i, 1) = oDrDatos!cCtaCod
        poControl.TextMatrix(i, 2) = oDrDatos!cUser
        poControl.TextMatrix(i, 3) = Format(oDrDatos!dFecha, "dd/MM/yyyy")
        poControl.TextMatrix(i, 4) = Format(oDrDatos!dFecha, "hh:mm:ss")
        poControl.TextMatrix(i, 5) = Trim(oDrDatos!cConvenioOrigen)
        If SSTab1.Tab = 1 Then
            poControl.TextMatrix(i, 6) = Trim(oDrDatos!cConvenioDestino)
        End If
        oDrDatos.MoveNext
    Next
End Sub

