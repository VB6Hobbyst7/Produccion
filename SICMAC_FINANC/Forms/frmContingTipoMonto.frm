VERSION 5.00
Begin VB.Form frmContingTipoMonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contingencias: Mantenimiento Tipo de Monto"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmContingTipoMonto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmContingTipoMonto.frx":030A
   ScaleHeight     =   3930
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin Sicmact.FlexEdit feTipoMonto 
      Height          =   1755
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3096
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Tipo de Monto"
      EncabezadosAnchos=   "400-1000-4300"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      BorderStyle     =   0
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registrar: "
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtTipoMontoDesc 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTipoMontoID 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmContingTipoMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim N As Integer
Dim oCon As DConecta
Dim oContingTM As DContingenciaTipoMonto

Private Sub cmdCancelar_Click()
    txtTipoMontoDesc.Text = ""
    lblTipoMontoID.Caption = ""
    txtTipoMontoDesc.SetFocus
    Frame1.Caption = "Registrar:"
    cmdGuardar.Caption = "Guardar"
End Sub

Private Sub cmdGuardar_Click()
  Dim desc As String
  Dim textoBoton As String
  textoBoton = cmdGuardar.Caption
  desc = txtTipoMontoDesc.Text
    Set oContingTM = New DContingenciaTipoMonto

   If textoBoton = "Guardar" Then
        If desc <> "" Then
            If MsgBox("Está seguro de registrar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            Call oContingTM.RegistrarTipoMontoPasivoContingente(desc)
            MsgBox "El Tipo Monto " & desc & " se ha registrado exitosamente", vbInformation, "Aviso"
            Call CargaDatos
        Else
            MsgBox "Falta ingresar la Descripcion", vbInformation, "Aviso"
        End If
    ElseIf textoBoton = "Modificar" Then
        If MsgBox("Está seguro de actualizar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Call oContingTM.ActualizarTipoMontoPasivoContingente(lblTipoMontoID.Caption, desc)
        MsgBox "Se ha actualizado exitosamente el tipo de monto", vbInformation, "Aviso"
        Call CargaDatos
        cmdGuardar.Caption = "Guardar"
        Frame1.Caption = "Registrar:"
    End If
    
    lblTipoMontoID.Caption = ""
    txtTipoMontoDesc.Text = ""
    txtTipoMontoDesc.SetFocus
End Sub

Private Sub feTipoMonto_OnRowChange(pnRow As Long, pnCol As Long)
    If feTipoMonto.Rows > 2 Then
        txtTipoMontoDesc.Text = ""
        lblTipoMontoID.Caption = ""
        Frame1.Caption = "Modificar:"
        lblTipoMontoID.Caption = feTipoMonto.TextMatrix(pnRow, 1)
        txtTipoMontoDesc.Text = feTipoMonto.TextMatrix(pnRow, 2)
        cmdGuardar.Caption = "Modificar"
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaDatos
End Sub

Private Sub CargaDatos()
    Dim rsListarTM As ADODB.Recordset
    Set oContingTM = New DContingenciaTipoMonto
    Dim i As Integer
    
    Set rsListarTM = oContingTM.ListarTipoMontoPasivoContingente
    Call LimpiaFlex(feTipoMonto)
    
    For i = 1 To rsListarTM.RecordCount
        feTipoMonto.AdicionaFila
        feTipoMonto.TextMatrix(i, 1) = rsListarTM!Codigo
        feTipoMonto.TextMatrix(i, 2) = rsListarTM!desc
        rsListarTM.MoveNext
    Next i
    
End Sub

