VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmCFTarifario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Carta Fianza Tarifario"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "frmCFTarifario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraControles 
      Height          =   5175
      Left            =   7320
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      Begin VB.TextBox TxtMontoMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   3930
         Width           =   1215
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4680
         Width           =   1215
      End
      Begin VB.ComboBox CboModalidad 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox TxtTasa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto Maximo"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   3690
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto Minimo"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   7095
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FraTarifario 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshTarifario 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
End
Attribute VB_Name = "frmCFTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim opt As String
Dim objPista As COMManejador.Pista
Dim loContFunct As COMNContabilidad.NCOMContFunciones

Sub ActDesControles(pbCboModalidad As Boolean, pbMoneda As Boolean, _
                    pbMonto As Boolean, pbTasa As Boolean, pbMontoMax As Boolean)

 CboModalidad.Enabled = pbCboModalidad
 CboMoneda.Enabled = pbMoneda
 TxtMonto.Enabled = pbMonto
 TxtTasa.Enabled = pbTasa
 TxtMontoMax.Enabled = pbMontoMax

End Sub

Private Sub CmdAceptar_Click()
Dim Op As Integer
Dim nCod As Integer
Dim lsMovNro As String 'MAVM 20100625 BAS II
If opt = "" Then Exit Sub
Dim CF As COMDCartaFianza.DCOMCartaFianza
Set CF = New COMDCartaFianza.DCOMCartaFianza
If Not IsNumeric(TxtMonto) Then
    MsgBox "Ingrese el Monto", vbCritical, "AVISO"
    Exit Sub
End If
If Not IsNumeric(TxtTasa) Then
    MsgBox "Ingrese una Tasa Correcta", vbCritical, "AVISO"
    Exit Sub
End If

lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing

Select Case opt
Case "M":
          Op = MsgBox("Esta seguro de Modificar", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            nCod = MshTarifario.TextMatrix(MshTarifario.Row, MshTarifario.Col)
            Call CF.ActualizaTarifario(val(TxtTasa.Text), val(TxtMonto.Text), nCod, val(TxtMontoMax.Text))
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Modificar Tarifario", nCod, gCodigo
          End If
          
Case "N":
         If CboMoneda.ListIndex = -1 Then
            MsgBox "Escoja una opcion", vbCritical, "AVISO"
            Exit Sub
         End If
        If CboModalidad.ListIndex = -1 Then
            MsgBox "Escoja una opcion", vbCritical, "AVISO"
            Exit Sub
         End If

          Op = MsgBox("Esta seguro de Grabar", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            'nCod = Val(Trim(Right(CboModalidad.Text, 3) & Right(CboMoneda.Text, 3)))
            If CF.ExisteTarifario(nCod) Then
                If MsgBox("Ya existe esa Tarifa parecida. Desea Continuar?", vbInformation + vbYesNo, "AVISO") = vbNo Then
                    Exit Sub
                End If
            End If
            Call CF.InsertaTarifario(val(TxtTasa.Text), val(TxtMonto.Text), nCod, val(Right(CboMoneda.Text, 3)), val(Right(CboModalidad.Text, 3)), val(TxtMontoMax.Text))
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Inserta Tarifario", nCod, gCodigo
          End If
End Select
Set CF = Nothing
opt = ""
CmdNuevo.Enabled = True
CmdModificar.Enabled = True
CmdSalir.Caption = "&Salir"
MshTarifario.Clear
MshTarifario.Rows = 2
CargaTarifario
Marco
End Sub

Private Sub CmdImprimir_Click()
Dim CF As COMNCartaFianza.NCOMCartaFianzaReporte 'NCartaFianzaReporte
Dim lsCad As String
Dim P As previo.clsprevio
Set CF = New COMNCartaFianza.NCOMCartaFianzaReporte 'NCartaFianzaReporte
Set P = New previo.clsprevio
lsCad = CF.nRepoTarifario(gImpresora)
P.Show Chr$(27) & Chr$(77) & lsCad, "Reportes de Creditos", True, , gImpresora

Set P = Nothing
Set CF = Nothing
End Sub

Private Sub CmdModificar_Click()
Call ActDesControles(False, False, True, True, True)
opt = "M"
CmdSalir.Caption = "&Cancelar"
CmdNuevo.Enabled = False
CmdModificar.Enabled = False
End Sub

Private Sub cmdNuevo_Click()
Call ActDesControles(True, True, True, True, True)
CmdModificar.Enabled = False
CmdSalir.Caption = "&Cancelar"
CmdNuevo.Enabled = False
opt = "N"
End Sub

Private Sub cmdsalir_Click()
If CmdSalir.Caption = "&Cancelar" Then
    CmdSalir.Caption = "&Salir"
    Call ActDesControles(False, False, False, False, False)
    CmdNuevo.Enabled = True
    CmdModificar.Enabled = True
    opt = ""
Else
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
CargaTarifario
Marco
Call CargaComboConstante(1011, CboMoneda)
Call CargaComboConstante(3402, CboModalidad)
Call ActDesControles(False, False, False, False, False)
TxtMonto = ""
TxtTasa = ""
opt = ""
Set objPista = New COMManejador.Pista
Set loContFunct = New COMNContabilidad.NCOMContFunciones
gsOpeCod = gCredMantTarfCF
End Sub

Sub CargaTarifario()
Dim rs As New ADODB.Recordset
Dim CF As COMDCartaFianza.DCOMCartaFianza
Set CF = New COMDCartaFianza.DCOMCartaFianza
Set rs = CF.RecuperaCF_Tarifario
'Set MshTarifario.DataSource = rs
With MshTarifario
While Not rs.EOF
    .TextMatrix(.Rows - 1, 0) = rs!cTarifCod
    .TextMatrix(.Rows - 1, 1) = rs!Modalidad & Space(50) & rs!nModalidad
    .TextMatrix(.Rows - 1, 2) = rs!Moneda & Space(50) & rs!nmoneda
    .TextMatrix(.Rows - 1, 3) = Format(rs!nTasaTrim, "0.00")
    .TextMatrix(.Rows - 1, 4) = Format(rs!nMontoMinimo, "0.00")
    .TextMatrix(.Rows - 1, 5) = Format(rs!nMontoMax, "0.00")
    .Rows = .Rows + 1
    rs.MoveNext
Wend
 .Rows = .Rows - 1
End With
Set CF = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set CF = Nothing
End Sub

Sub Marco()
With MshTarifario
    .TextMatrix(0, 0) = " Cod."
    .TextMatrix(0, 1) = " Modalidad"
    .TextMatrix(0, 2) = " Moneda"
    .TextMatrix(0, 3) = " Tasa"
    .TextMatrix(0, 4) = " Monto Min."
    .TextMatrix(0, 5) = " Monto Max."
    .ColWidth(0) = 500
    .ColWidth(1) = 3500
    .ColWidth(2) = 1000
    .ColWidth(3) = 700
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
End With

End Sub

Private Sub MshTarifario_Click()
If opt = "M" Then
  With MshTarifario
    If .Row > 0 Then
    CboModalidad.ListIndex = CInt(Right(.TextMatrix(.Row, 1), 2)) - 1
    CboMoneda.ListIndex = CInt(Right(.TextMatrix(.Row, 2), 2)) - 1
    TxtMonto = Format(.TextMatrix(.Row, 4), "0.00")
    TxtTasa = Format(.TextMatrix(.Row, 3), "0.00")
    End If
  End With
End If
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoMax.SetFocus
    End If

End Sub

Private Sub TxtMontoMax_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Me.CmdAceptar.SetFocus
    End If

End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtMonto.SetFocus
    End If
End Sub
