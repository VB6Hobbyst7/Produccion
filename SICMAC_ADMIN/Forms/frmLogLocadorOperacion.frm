VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogLocadorOperacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   1515
   ClientTop       =   1860
   ClientWidth     =   8790
   Icon            =   "frmLogLocadorOperacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Contrato "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   120
      TabIndex        =   10
      Top             =   1820
      Width           =   3735
      Begin VB.TextBox txtNroContrato 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vigencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   3960
      TabIndex        =   5
      Top             =   1820
      Width           =   4695
      Begin VB.TextBox txtFecIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtFecFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2700
         TabIndex        =   8
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Identificación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8535
      Begin VB.TextBox txtFuncion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1260
         Width           =   7155
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   7155
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   660
         Width           =   7155
      End
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2475
         TabIndex        =   3
         Top             =   320
         Width           =   390
      End
      Begin VB.TextBox txtPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1730
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   315
         Left            =   2895
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   5385
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Función"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame fraReg 
      Height          =   2655
      Left            =   120
      TabIndex        =   19
      Top             =   2700
      Width           =   8535
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7020
         TabIndex        =   27
         Top             =   2220
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   1875
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3307
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.CommandButton cmdProvPago 
         Caption         =   "Autorizar el Pago"
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   2220
         Width           =   1695
      End
   End
   Begin VB.Frame fraProg 
      Height          =   2595
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   6045
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   3480
         TabIndex        =   29
         Top             =   1440
         Width           =   4035
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         ItemData        =   "frmLogLocadorOperacion.frx":08CA
         Left            =   3480
         List            =   "frmLogLocadorOperacion.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1020
         Width           =   4035
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   4980
         TabIndex        =   24
         Top             =   2040
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6300
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento de pago"
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         Top             =   1500
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento de pago"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmLogLocadorOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String
Dim nRegLoc As Integer
Dim sSQL As String, nMontoContrato As Currency, nSumaCuotas As Currency, nSumaPagada As Currency

Public Sub Inicio(ByVal psOpeCod As String)
cOpeCod = psOpeCod
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
fraReg.Visible = True
fraProg.Visible = False
End Sub

Private Sub cmdGrabar_Click()
fraReg.Visible = True
fraProg.Visible = False
End Sub

Private Sub cmdProvPago_Click()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim sSQL As String

txtDescripcion.Text = "CUOTA Nº " + MSFlex.TextMatrix(MSFlex.row, 1) + " PARA CANCELAR EL DIA " + MSFlex.TextMatrix(MSFlex.row, 2) + " POR " + MSFlex.TextMatrix(MSFlex.row, 3) + " " + MSFlex.TextMatrix(MSFlex.row, 4)

cboDoc.Clear
sSQL = "select o.nDocTpo,d.cDocDesc from OpeDoc o inner join Documento d on o.nDocTpo = d.nDocTpo where cOpeCod = '" & cOpeCod & "'"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboDoc.AddItem rs!cDocDesc
         cboDoc.ItemData(cboDoc.ListCount - 1) = rs!nDocTpo
         rs.MoveNext
      Loop
      cboDoc.ListIndex = 0
   End If
End If
txtNroDoc.Text = ""
fraReg.Visible = False
fraProg.Visible = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
FlexCuotas
cmdProvPago.Enabled = False
Select Case cOpeCod
    Case "541020"
         'cmdOperacion.Caption = "Solicitar Pago"
    Case "541030"
    
End Select
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersCod = X.sPersCod
   VerDatos txtPersCod
End If
End Sub

Private Sub VerDatos(ByVal vpPersCod As String)
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer, k As Integer
Dim nCuota As Integer

nMontoContrato = 0
nSumaCuotas = 0
nSumaPagada = 0
FlexCuotas

If oConn.AbreConexion Then

   sSQL = "select x.nRegLocador, x.cPersCod, cPersona=replace(p.cPersNombre,'/',' '),cNroContrato,dFechaInicio,dFechaTermino, " & _
   "      aa.cAreaDescripcion as cArea, a.cAgeDescripcion as cAgencia,x.nMontoContrato, " & _
   "      t.cConsDescripcion as cTipoContrato, f.cConsDescripcion as cFuncion  " & _
   " from LogLocadores x inner join Persona p on x.cPersCod=p.cPersCod " & _
   "                     inner join Agencias a on x.cAgeCod = a.cAgeCod " & _
   "                     inner join Areas aa on x.cAreaCod = aa.cAreaCod " & _
   "                     inner join (select nConsValor,cConsDescripcion from Constante where nConsCod =9131 and nConsCod<>nConsValor) t on x.nTipoContrato = t.nConsValor " & _
   "                     inner join (select nConsValor,cConsDescripcion from Constante where nConsCod =9132 and nConsCod<>nConsValor) f on x.nFuncionCod = f.nConsValor " & _
   " where x.cPersCod = '" & vpPersCod & "'"

   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      nRegLoc = rs!nRegLocador
      txtPersCod.Text = rs!cPersCod
      txtPersona.Text = rs!cPersona
      txtNroContrato.Text = rs!cNroContrato
      'txtMonto.Text = FNumero(rs!nMontoContrato)
      txtFecIni = rs!dFechaInicio
      txtFecFin = rs!dFechaTermino
      txtFuncion.Text = rs!cFuncion
      txtAgencia.Text = rs!cAgencia
      txtArea.Text = rs!cArea
      nMontoContrato = rs!nMontoContrato
      cmdProvPago.Enabled = True
   Else
      cmdProvPago.Enabled = False
   End If
   
   i = 0
   k = 0
   nSumaCuotas = 0
   nSumaPagada = 0
   
   sSQL = "Select * from LogLocadoresCronograma where nRegLocador = " & nRegLoc & " "
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = Format(rs!nNroCuota, "00")
         MSFlex.TextMatrix(i, 2) = rs!dFechaPago
         Select Case rs!nMoneda
             Case 0
                  MSFlex.TextMatrix(i, 3) = ""
             Case 1
                  MSFlex.TextMatrix(i, 3) = "S/."
             Case 2
                  MSFlex.TextMatrix(i, 3) = "US$"
         End Select
         
         MSFlex.TextMatrix(i, 4) = FNumero(rs!nMontoAPagar)
         MSFlex.TextMatrix(i, 5) = FNumero(rs!nMontoPagado)
         
         nSumaCuotas = nSumaCuotas + rs!nMontoAPagar
         If rs!nEstadoCuota = 2 Then
            nSumaPagada = nSumaPagada + rs!nMontoPagado
            MSFlex.TextMatrix(i, 6) = "Cancelada"
         Else
            MSFlex.TextMatrix(i, 6) = "Pendiente"
         End If
         rs.MoveNext
      Loop
   End If
End If
End Sub



         'MSFlex.TextMatrix(i, 6) = FNumero(rs!nMontoPagado)
         'MSFlex.row = i
         'MSFlex.Col = 1
         'MSFlex.CellPictureAlignment = 4
         'MSFlex.CellForeColor = "&H80000005"
         'If rs!nMontoPagado > 0 Then
         '   Set MSFlex.CellPicture = imgCheck(2)
         'Else
         '   Set MSFlex.CellPicture = imgCheck(0)
         'End If

'         If rs!nEstadoCuota = 0 Then
'            i = i + 1
'            lsvReg.ListItems.Add i
'            lsvReg.ListItems(i).SubItems(1) = rs!nNroCuota
'            lsvReg.ListItems(i).SubItems(2) = rs!dFechaPago
'            lsvReg.ListItems(i).SubItems(3) = FNumero(rs!nMontoAPagar)
'            nSumaCuotas = nSumaCuotas + rs!nMontoPagado
'         End If
'
'         If rs!nEstadoCuota = 1 Then
'            k = k + 1
'            lsvReg.ListItems.Add k
'            lsvReg.ListItems(k).SubItems(1) = rs!nNroCuota
'            lsvReg.ListItems(k).SubItems(2) = rs!dFechaPago
'            lsvReg.ListItems(k).SubItems(3) = FNumero(rs!nMontoPagado)
'         End If
'




'Private Sub MSFlex_Click()
'Dim nFil As Integer, nCol As Integer
'
'nFil = MSFlex.row
'If MSFlex.Col = 1 Then
'   MSFlex.row = nFil
'   MSFlex.Col = 1
'
'   If Len(MSFlex.TextMatrix(nFil, 1)) > 0 Then
'      Set MSFlex.CellPicture = imgCheck(0)
'      MSFlex.TextMatrix(nFil, 1) = ""
'   Else
'      Set MSFlex.CellPicture = imgCheck(1)
'      MSFlex.TextMatrix(nFil, 1) = "."
'   End If
'
'End If
'End Sub

'Private Sub MSFlex_DblClick()
'Dim i As Integer
'i = MSFlex.row
'
'End Sub

Sub FlexCuotas()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 300
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 400:  MSFlex.TextMatrix(0, 1) = " Nº":         MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 1100: MSFlex.TextMatrix(0, 2) = "Fecha Pago":  MSFlex.ColAlignment(2) = 4
MSFlex.ColWidth(3) = 750: MSFlex.TextMatrix(0, 3) = "Moneda"
MSFlex.ColWidth(4) = 1200: MSFlex.TextMatrix(0, 4) = "Monto a Pagar"
MSFlex.ColWidth(5) = 1200: MSFlex.TextMatrix(0, 5) = "Monto Pagado"
MSFlex.ColWidth(6) = 3000: MSFlex.TextMatrix(0, 6) = "Estado"

'MSFlex.ColWidth(4) = 1200: MSFlex.TextMatrix(0, 4) = " Monto a Pagar"
'MSFlex.ColWidth(5) = 1200: MSFlex.TextMatrix(0, 5) = " Monto Pagado"
'MSFlex.ColWidth(6) = 2000: MSFlex.TextMatrix(0, 6) = " Estado"
End Sub

