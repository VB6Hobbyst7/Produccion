VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogEnviosTarifas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3660
   ClientLeft      =   1710
   ClientTop       =   2790
   ClientWidth     =   8070
   Icon            =   "frmLogEnviosTarifas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8070
   Begin VB.Frame fraVis 
      Caption         =   "Tarifas para envío de Correspondencia "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   120
      TabIndex        =   20
      Top             =   60
      Width           =   7815
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1500
         TabIndex        =   24
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6360
         TabIndex        =   23
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   3000
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2655
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraReg 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   7815
         Begin VB.TextBox txtUbigeoOrigen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   360
            Width           =   4905
         End
         Begin VB.CommandButton cmdUbigeoOrig 
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
            Height          =   270
            Left            =   2370
            TabIndex        =   14
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtUbigeoDestino 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   720
            Width           =   4905
         End
         Begin VB.CommandButton cmdUbigeoDest 
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
            Height          =   270
            Left            =   2370
            TabIndex        =   12
            Top             =   750
            Width           =   315
         End
         Begin VB.TextBox txtUbigeoOrig 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   1755
         End
         Begin VB.TextBox txtUbigeoDest 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Origen"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Destino"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   780
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   0
         TabIndex        =   8
         Top             =   2100
         Width           =   7815
         Begin VB.TextBox txtTarifaMonto 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Costo"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Frame Frame3 
         Height          =   795
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   7815
         Begin VB.TextBox txtTipoEnvio 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   4905
         End
         Begin VB.CommandButton cmdTipoEnvio 
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
            Height          =   270
            Left            =   2370
            TabIndex        =   4
            Top             =   330
            Width           =   315
         End
         Begin VB.TextBox txtTipoCod 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Envío"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6540
         TabIndex        =   1
         Top             =   3000
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmLogEnviosTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgregar_Click()
fraVis.Visible = False
fraReg.Visible = True
End Sub

Private Sub cmdCancelar_Click()
fraReg.Visible = False
fraVis.Visible = True
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim sSQL As String, rs As New ADODB.Recordset
Dim nTipoEnvio As Integer, cUbiOrig As String, cUbiDest As String

If VNumero(txtTipoCod) <= 0 Then
   MsgBox "El tipo de envío no es válido..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If Len(Trim(txtUbigeoOrig.Text)) < 10 Or Len(Trim(txtUbigeoDest.Text)) < 10 Then
   MsgBox "Una ubicación geográfica no es válida..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If VNumero(txtTarifaMonto.Text) <= 0 Then
   MsgBox "Monto de tarifa no válido..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

sSQL = "Select * from LogCtrlEnviosTarifas where nTipoEnvio = " & CInt(txtTipoCod) & " and " & _
       "         left(cUbigeoOrigen,5) = '" & Left(txtUbigeoOrig.Text, 5) & "' and left(cUbigeoDestino,5) = '" & Left(txtUbigeoDest.Text, 5) & "' " & _
       "         "
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      MsgBox "Ya existe un registro para el Origen/Destino a nivel de Provincias..." + Space(10), vbInformation, "Aviso"
      Exit Sub
   End If
   oConn.CierraConexion
End If

If MsgBox("¿ Está seguro de grabar la Tarifa indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
  
   If oConn.AbreConexion Then
      sSQL = " INSERT INTO LogCtrlEnviosTarifas (nTipoEnvio, cUbigeoOrigen, cUbigeoDestino, nTarifaMonto ) " & _
           " VALUES (" & CInt(txtTipoCod) & ",'" & txtUbigeoOrig.Text & "','" & txtUbigeoDest.Text & "'," & VNumero(txtTarifaMonto.Text) & ") " & _
           " "
      oConn.Ejecutar sSQL
      
      sSQL = " INSERT INTO LogCtrlEnviosTarifas (nTipoEnvio, cUbigeoOrigen, cUbigeoDestino, nTarifaMonto ) " & _
           " VALUES (" & CInt(txtTipoCod) & ",'" & txtUbigeoDest.Text & "','" & txtUbigeoOrig.Text & "'," & VNumero(txtTarifaMonto.Text) & ") " & _
           " "
      oConn.Ejecutar sSQL
      oConn.CierraConexion
   End If
   
   fraReg.Visible = False
   fraVis.Visible = True
End If
End Sub

Private Sub cmdTipoEnvio_Click()
Dim sSQL As String
sSQL = "Select nConsValor, cConsDescripcion from Constante where nConsCod = 9135 and nConsCod <> nConsValor"
frmLogSelector.Consulta sSQL, "Seleccione Tipo de Envío"
If frmLogSelector.vpHaySeleccion Then
   txtTipoCod.Text = frmLogSelector.vpCodigo
   txtTipoEnvio.Text = frmLogSelector.vpDescripcion
End If
End Sub

Private Sub cmdUbigeoDest_Click()
frmLogProSelSeleUbiGeo.Show 1
If Len(Trim(frmLogProSelSeleUbiGeo.gvCodigo)) > 0 Then
   txtUbigeoDest.Text = frmLogProSelSeleUbiGeo.gvCodigo
   txtUbigeoDestino.Text = frmLogProSelSeleUbiGeo.gvNoddo
End If
End Sub

Private Sub cmdUbigeoOrig_Click()
frmLogProSelSeleUbiGeo.Show 1
If Len(Trim(frmLogProSelSeleUbiGeo.gvCodigo)) > 0 Then
   txtUbigeoOrig.Text = frmLogProSelSeleUbiGeo.gvCodigo
   txtUbigeoOrigen.Text = frmLogProSelSeleUbiGeo.gvNoddo
End If
End Sub

Private Sub Form_Load()
CentraForm Me
txtTarifaMonto.Locked = False
ListaTarifasEnvio
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub ListaTarifasEnvio()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String, i As Integer

FlexLista

sSQL = "select t.nTipoEnvio, e.cTipoEnvio,o.cUbigeoDescripcion as cUbigeoOrigen, d.cUbigeoDescripcion as cUbigeoDestino,t.nTarifaMonto " & _
"   from LogCtrlEnviosTarifas t inner join (select nConsValor as nTipoEnvio, cConsDescripcion as cTipoEnvio from Constante where nConsCod = 9135 and nConsCod<>nConsValor) e on t.nTipoEnvio = e.nTipoEnvio " & _
"       left join UbicacionGeografica o on t.cUbigeoOrigen = o.cUbigeoCod " & _
"       left join UbicacionGeografica d on t.cUbigeoDestino = d.cUbigeoCod " & _
" WHERE t.nActivo = 1 " & _
" order by t.nTipoEnvio"
 
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cTipoEnvio
         MSFlex.TextMatrix(i, 2) = rs!cUbigeoOrigen
         MSFlex.TextMatrix(i, 3) = rs!cUbigeoDestino
         MSFlex.TextMatrix(i, 4) = FNumero(rs!nTarifaMonto)
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FlexLista()
MSFlex.Clear
MSFlex.RowHeight(0) = 300
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 1900: MSFlex.TextMatrix(0, 1) = "Tipo de Envío"
MSFlex.ColWidth(2) = 2000: MSFlex.TextMatrix(0, 2) = "Desde (Origen)         --->"
MSFlex.ColWidth(3) = 2000: MSFlex.TextMatrix(0, 3) = "a (Destino)"
MSFlex.ColWidth(4) = 1250: MSFlex.TextMatrix(0, 4) = "Costo de envío"
MSFlex.ColWidth(5) = 0
End Sub

Private Sub txtTarifaMonto_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumDec(txtTarifaMonto, KeyAscii)
If nKeyAscii = 13 Then
   cmdGrabar.SetFocus
End If
End Sub
