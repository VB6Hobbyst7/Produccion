VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCredPersEstado 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   2370
   ClientTop       =   2250
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "frmCredPersEstado.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6135
      TabIndex        =   3
      Top             =   2850
      Width           =   1140
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4980
      TabIndex        =   2
      Top             =   2850
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   7230
      Begin MSDataGridLib.DataGrid DGPersonas 
         Height          =   2475
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   4366
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cCtaCod"
            Caption         =   "CREDITO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cPersNombre"
            Caption         =   "NOMBRE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               DividerStyle    =   1
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4289.953
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCredPersEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RGrid As ADODB.Recordset
Dim vsSelecPers As String
Dim vsSelecDacion As Long
Dim vTipoBusqueda As Long
Dim vAgenciaSelect As String
Public vcodper As String

Private Sub CargaGrid(ByVal pnTipoCredPers As Variant, Optional ByVal pMatProd As Variant = Nothing, _
    Optional ByVal pbRefin As Boolean = False, Optional pbMuestraTodos As Boolean = False)
Dim oCreditos As COMDCredito.DCOMCreditos
    On Error GoTo ERRORCargaGrid
    Set oCreditos = New COMDCredito.DCOMCreditos
    Set RGrid = oCreditos.RecuperaPersonasEstadoCred(pnTipoCredPers, pMatProd, pbRefin, pbMuestraTodos)
    Set DGPersonas.DataSource = RGrid
    DGPersonas.Refresh
    Set oCreditos = Nothing
    Exit Sub
    
ERRORCargaGrid:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub CargaGridAgencia(ByVal pnTipoCredPers As Variant, _
    Optional ByVal pMatProd As Variant = Nothing, _
    Optional ByVal pbRefin As Boolean = False, _
    Optional pbMuestraTodos As Boolean = False, _
    Optional ByVal psAgencia As String, _
    Optional pbLeasing As Boolean = False, _
    Optional ByVal pbMicroMulti As Boolean = False, _
    Optional pbInfoGas As Boolean = False, _
    Optional pbAmpliado As Boolean = False, _
    Optional pgsCodCargo As String = "")
    'WIOR 20120518 Se agrego el parametro pbMicroMulti
    'BRGO 20120707 Se agregò paràmetro pbInfoGas
    'LUCV20180417, Agregó pbAmpliado: según incidente en la edición de créditos ampliados
    
    Dim oCreditos As COMDCredito.DCOMCreditos
    On Error GoTo ERRORCargaGrid
    Set oCreditos = New COMDCredito.DCOMCreditos
    'Set RGrid = oCreditos.RecuperaPersonasEstadoCred(pnTipoCredPers, pMatProd, pbRefin, pbMuestraTodos)
'    Set RGrid = New Recordset
    'WIOR 20120518 Se agrego el parametro pbMicroMulti
    'LUCV20180417, Agregó pbAmpliado
    Set RGrid = oCreditos.RecuperaPersonasEstadoCredAgencia(pnTipoCredPers, psAgencia, pMatProd, pbRefin, pbMuestraTodos, pbLeasing, pbMicroMulti, pbInfoGas, pbAmpliado, pgsCodCargo)
    
    Set DGPersonas.DataSource = RGrid
    DGPersonas.Refresh
    Set oCreditos = Nothing
    Exit Sub
    
ERRORCargaGrid:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub CargaGridDaciones()
Dim oDCred As COMDCredito.DCOMCredito

    Set oDCred = New COMDCredito.DCOMCredito
    Set RGrid = New ADODB.Recordset
    Set RGrid = oDCred.RecuperaDacionesPago
    Set DGPersonas.DataSource = RGrid
    Set oDCred = Nothing
    
End Sub

Public Function BuscaDacionesPago(ByVal psCaption As String) As Long
    vTipoBusqueda = 2
    Me.Caption = psCaption
    DGPersonas.Columns(0).DataField = "nNroGarantRec"
    DGPersonas.Columns(0).Width = "500"
    DGPersonas.Columns(1).DataField = "cPersNombre"
    DGPersonas.Columns.Add 2
    DGPersonas.Columns(2).DataField = "cCtaCod"
    
    vsSelecDacion = -1
    Call CargaGridDaciones
    Screen.MousePointer = 0
    Me.Show 1
    Set RGrid = Nothing
    
    BuscaDacionesPago = vsSelecDacion
End Function

Public Function Inicio(ByVal pnTipoCredPers As Variant, ByVal psCaption As String, _
    Optional ByVal pMatProd As Variant = Nothing, _
    Optional ByVal pbRefin As Boolean = False, _
    Optional ByVal pbMuestraTodos As Boolean = False, _
    Optional ByVal psAgencia As String = "ALL", _
    Optional pbLeasing As Boolean = False, _
    Optional ByVal pbMicroMulti As Boolean = False, _
    Optional pbInfoGas As Boolean = False, _
    Optional pbAmpliado As Boolean = False, _
    Optional pgsCodCargo As String = "") As String
    'WIOR 20120518 Se agrego el parametro pbMicroMulti
    'LUCV20180417, Agregó pbAmpliado: según incidente en la edición de créditos ampliados
    'JOEP20190205 CP pgsCodCargo
Dim i As Integer
   
    vTipoBusqueda = 1
    Me.Caption = psCaption
    vsSelecPers = ""
    vAgenciaSelect = IIf(psAgencia = "ALL", "", psAgencia)
    'WIOR 20120518 Se agrego el parametro pbMicroMulti
    'LUCV20180417, Agregó pbAmpliado
    'JOEP20190205 CP pgsCodCargo
    Call CargaGridAgencia(pnTipoCredPers, pMatProd, pbRefin, pbMuestraTodos, psAgencia, pbLeasing, pbMicroMulti, pbInfoGas, pbAmpliado, pgsCodCargo)
    
    Screen.MousePointer = 0
    Me.Show 1
    Set RGrid = Nothing
    
    Inicio = vsSelecPers
   
End Function

Private Sub CmdAceptar_Click()
    If RGrid Is Nothing Then
        Unload Me
        Exit Sub
    End If
    If RGrid.RecordCount > 0 Then
        Select Case vTipoBusqueda
            Case 1
                vsSelecPers = RGrid.Fields(2)
            Case 2
                vsSelecDacion = RGrid.Fields(0)
        End Select
         vcodper = RGrid.Fields(0)
    Else
        Select Case vTipoBusqueda
            Case 1
                vsSelecPers = ""
            Case 2
                vsSelecDacion = -1
        End Select
        
    End If
    Unload Me
End Sub


Private Sub cmdCancelar_Click()
    Select Case vTipoBusqueda
        Case 1
            vsSelecPers = ""
        Case 2
            vsSelecDacion = -1
    End Select
    Unload Me
End Sub

Private Sub DGPersonas_KeyPress(KeyAscii As Integer)
    Dim rs As ADODB.Recordset
    Dim nPos As Integer

    Set rs = RGrid.Clone

    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        RGrid.MoveFirst
    End If

    nPos = 0
'    Set DGPersonas.DataSource = RGrid
'    DGPersonas.Refresh
'    DGPersonas.Row = 0
    Do Until rs.EOF
        nPos = nPos + 1
        If Mid(rs!cPersNombre, 1, 1) = UCase(Chr(KeyAscii)) Then
            RGrid.Move nPos - 1
            'DGPersonas.Row = nPos - 1
            'DGPersonas.Refresh
            Exit Do
        End If
        rs.MoveNext
    Loop

    Set rs = Nothing

' RGrid.Find "cPersNombre like '" & Chr(KeyAscii) & "'"
' Set DGPersonas.DataSource = RGrid
' DGPersonas.Refresh

End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    CmdAceptar.Default = True
End Sub

'WIOR 20160620 ***
Public Function InicioDesBloqSobreEnd(ByVal psCaption As String, ByVal psCargo As String) As String
Dim oDCred As COMDCredito.DCOMCredito

    vTipoBusqueda = 1
    Me.Caption = psCaption
    vsSelecPers = ""

    Set oDCred = New COMDCredito.DCOMCredito
    Set RGrid = oDCred.SobreEndCreditosADesbloq(psCargo)
    Set DGPersonas.DataSource = RGrid
    Set oDCred = Nothing
    
    Screen.MousePointer = 0
    Me.Show 1
    
    InicioDesBloqSobreEnd = vsSelecPers
   
End Function
'WIOR FIN ********


