VERSION 5.00
Begin VB.Form frmCredRiesgoLimiteProductoAutoriza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud fuera de Limite por Tipo de Producto"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmCredRiesgoLimiteProductoAutoriza.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbAgencia 
      Height          =   315
      ItemData        =   "frmCredRiesgoLimiteProductoAutoriza.frx":030A
      Left            =   960
      List            =   "frmCredRiesgoLimiteProductoAutoriza.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   3735
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Height          =   360
      Left            =   8280
      TabIndex        =   3
      Top             =   5340
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   9480
      TabIndex        =   2
      Top             =   5340
      Width           =   1050
   End
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   5340
      Width           =   6135
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Height          =   360
      Left            =   7080
      TabIndex        =   0
      Top             =   5340
      Width           =   1050
   End
   Begin SICMACT.FlexEdit feCredInfRiesgo 
      Height          =   4740
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   8361
      Cols0           =   13
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Crédito-Titular-Tipo Producto-Moneda-Monto-Monto MN-Credito MN-Limite-% con este Credito-Credito.-nCodAge"
      EncabezadosAnchos=   "400-2000-1800-2500-2200-1000-1200-1800-1800-1500-1800-2000-0"
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
      ColumnasAEditar =   "X-1-2-3-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-L-R-C-C-R-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-1-0-0"
      CantEntero      =   10
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblAgencia 
      Caption         =   "Agencia :"
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
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5340
      Width           =   615
   End
End
Attribute VB_Name = "frmCredRiesgoLimiteProductoAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDPersGen As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim rsDatosAgeTotal As ADODB.Recordset
Dim nValorInicio As Integer

Private Sub cmbAgencia_Click()
Dim nValorTipAge As String
Dim i As Integer

If nValorInicio <> 0 Then

        nValorTipAge = (cmbAgencia.ItemData(cmbAgencia.ListIndex))
        nValorTipAge = IIf(nValorTipAge <= 9, "0" & nValorTipAge, nValorTipAge)
  
    If nValorTipAge <> 0 Then
  
        rsDatosAgeTotal.Filter = " cAgeCod LIKE '*" + nValorTipAge + "*'"
 
        feCredInfRiesgo.Clear
        feCredInfRiesgo.FormaCabecera
        Call LimpiaFlex(feCredInfRiesgo)
        For i = 1 To rsDatosAgeTotal.RecordCount
            feCredInfRiesgo.AdicionaFila
                feCredInfRiesgo.TextMatrix(i, 1) = rsDatosAgeTotal!cAgeDescripcion
                feCredInfRiesgo.TextMatrix(i, 2) = rsDatosAgeTotal!cCtaCod
                feCredInfRiesgo.TextMatrix(i, 3) = rsDatosAgeTotal!cPersNombre
                feCredInfRiesgo.TextMatrix(i, 4) = rsDatosAgeTotal!cTpoProdDesc
                feCredInfRiesgo.TextMatrix(i, 5) = rsDatosAgeTotal!cMoneda
                feCredInfRiesgo.TextMatrix(i, 6) = Format(rsDatosAgeTotal!nMontoSol, "#,##0.00")
                feCredInfRiesgo.TextMatrix(i, 7) = Format(rsDatosAgeTotal!nMontoMN, "#,##0.00")
                feCredInfRiesgo.TextMatrix(i, 8) = Format(rsDatosAgeTotal!nZonaMN, "#,##0.00")
                feCredInfRiesgo.TextMatrix(i, 9) = Format(rsDatosAgeTotal!nLimite, "#,##0.00")
                feCredInfRiesgo.TextMatrix(i, 10) = Format(rsDatosAgeTotal!nPorcConEsteCred, "#,##0.00")
                feCredInfRiesgo.TextMatrix(i, 11) = rsDatosAgeTotal!cZonaDesc
                feCredInfRiesgo.TextMatrix(i, 12) = rsDatosAgeTotal!cAgeCod
                rsDatosAgeTotal.MoveNext
        Next i
    Else
    Call CargarSolicitudes
    End If
End If
 nValorInicio = 1
End Sub

Private Sub cmdAutorizar_Click()
AutorizarRechazarSolicitud (1)
End Sub

Private Sub cmdRechazar_Click()
AutorizarRechazarSolicitud (2)
End Sub

Private Sub AutorizarRechazarSolicitud(ByVal pnEstado As Integer)

Set oDPersGen = New COMDCredito.DCOMCredito

If Trim(txtGlosa.Text) = "" Then
    MsgBox "Debe ingresar la glosa", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If

If feCredInfRiesgo.TextMatrix(1, 2) <> "" Then
        If pnEstado = 1 Then
            If MsgBox("Esta opción autorizará sugerir un crédito que superará los límites por Zona Geografica, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Else
            If MsgBox("Esta opción rechazará la solicitud de autorización del crédito, ésto podrá ser sugerido, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If

        Call oDPersGen.ActualizarSolicitudAutorizacionTpCredito(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2), Trim(Me.txtGlosa.Text), pnEstado)
Else
    MsgBox "Favor de seleccionar el Credito.", vbInformation, "Aviso"
    Exit Sub
End If

Set oDPersGen = Nothing
MsgBox "La solicitud fue " & IIf(pnEstado = 1, "autorizada", "rechazada"), vbInformation, "Aviso"
txtGlosa.Text = ""
CargarSolicitudes
End Sub


Private Sub cmdSalir_Click()
    Unload Me
nValorInicio = 0
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    nValorInicio = 0
    CargarSolicitudes
End Sub

Private Sub CargarSolicitudes()
Dim i As Integer
Dim rsCombo As ADODB.Recordset
Dim objAg As COMDConstantes.DCOMAgencias
Set objAg = New COMDConstantes.DCOMAgencias
Set oDPersGen = New COMDCredito.DCOMCredito
Set rs = oDPersGen.RecuperaSolicitudTpCredito
 
Set rsDatosAgeTotal = rs.Clone

    feCredInfRiesgo.Clear
    feCredInfRiesgo.FormaCabecera
    Call LimpiaFlex(feCredInfRiesgo)
    For i = 1 To rs.RecordCount
        feCredInfRiesgo.AdicionaFila
            feCredInfRiesgo.TextMatrix(i, 1) = rs!cAgeDescripcion
            feCredInfRiesgo.TextMatrix(i, 2) = rs!cCtaCod
            feCredInfRiesgo.TextMatrix(i, 3) = rs!cPersNombre
            feCredInfRiesgo.TextMatrix(i, 4) = rs!cTpoProdDesc
            feCredInfRiesgo.TextMatrix(i, 5) = rs!cMoneda
            feCredInfRiesgo.TextMatrix(i, 6) = Format(rs!nMontoSol, "#,##0.00")
            feCredInfRiesgo.TextMatrix(i, 7) = Format(rs!nMontoMN, "#,##0.00")
            feCredInfRiesgo.TextMatrix(i, 8) = Format(rs!nZonaMN, "#,##0.00")
            feCredInfRiesgo.TextMatrix(i, 9) = Format(rs!nLimite, "#,##0.00")
            feCredInfRiesgo.TextMatrix(i, 10) = Format(rs!nPorcConEsteCred, "#,##0.00")
            feCredInfRiesgo.TextMatrix(i, 11) = rs!cZonaDesc
            feCredInfRiesgo.TextMatrix(i, 12) = rs!cAgeCod
            rs.MoveNext
    Next i

If nValorInicio = 0 Then
    Set rsCombo = objAg.ObtieneAgencias
    cmbAgencia.AddItem "Todos"
        Do Until rsCombo.EOF
            cmbAgencia.AddItem "" & rsCombo!cConsDescripcion
            cmbAgencia.ItemData(cmbAgencia.NewIndex) = "" & rsCombo!nConsValor
            rsCombo.MoveNext
        Loop
    cmbAgencia.ListIndex = 0
End If

Set oDPersGen = Nothing
RSClose rsCombo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nValorInicio = 0
End Sub


