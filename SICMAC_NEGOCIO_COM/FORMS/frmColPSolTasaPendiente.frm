VERSION 5.00
Begin VB.Form frmColPSolTasaPendiente 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   7050
   ClientTop       =   4545
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
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
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Width           =   1140
   End
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
      Left            =   6600
      TabIndex        =   0
      Top             =   2640
      Width           =   1140
   End
   Begin SICMACT.FlexEdit FECredito 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4683
      Cols0           =   4
      ScrollBars      =   2
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "N°-Cuenta-Persona-Estado"
      EncabezadosAnchos=   "400-2000-4000-1000"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "N°"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmColPSolTasaPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Descripción : Formulario donde se listan las cuentas que tienen solicitud de tasa preferencial
'** Creación    : MACM 22-03-2021
'**********************************************************************************************
Option Explicit
Dim lsCtaCod As String
Dim RGrid As ADODB.Recordset

Public Sub cargarDatos()
    Dim lnPosI As Long
    Dim oCreditos As COMDCredito.DCOMNivelAprobacion
    Set oCreditos = New COMDCredito.DCOMNivelAprobacion
    On Error GoTo ERRORCargaGrid
    
    Set RGrid = oCreditos.RecuperaSolPendientesAge(gsCodAge)
    
    If RGrid.EOF And RGrid.BOF Then
        Me.CmdAceptar.Enabled = False
    Else
        Set FECredito.Recordset = RGrid
        Me.CmdAceptar.Enabled = True
    End If
    
    'Set DGPersonas.DataSource = RGrid
    
    
    'lnPosI = 0
    'For i = 1 To FECredito.rows - 1
     '   If FECredito.TextMatrix(i, 1) = "." Then
            
     '   End If
        
        
    Set oCreditos = Nothing
    Exit Sub
    
ERRORCargaGrid:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Sub Inicio(ByRef pcCtaCod As String)
    pcCtaCod = ""
    Call cargarDatos
    Me.Show 1
    
    pcCtaCod = lsCtaCod
End Sub

Private Sub cmdAceptar_Click()
Dim lnCuenta As String
Dim lrCredPig As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigJoyasDet As ADODB.Recordset
    If Len(Trim(FECredito.TextMatrix(FECredito.row, 1))) <> "18" Then
        MsgBox "Por Favor, debe seleccionar el crédito", vbInformation, "Aviso!"
    Else
        lsCtaCod = Trim(FECredito.TextMatrix(FECredito.row, 1))
        Unload Me
    End If
    End Sub

Private Sub cmdCancelar_Click()
    lsCtaCod = ""
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
    Do Until rs.EOF
        nPos = nPos + 1
        If Mid(rs!cPersNombre, 1, 1) = UCase(Chr(KeyAscii)) Then
            RGrid.Move nPos - 1
            Exit Do
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

