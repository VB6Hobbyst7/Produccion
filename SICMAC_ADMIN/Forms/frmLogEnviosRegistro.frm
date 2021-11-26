VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogEnviosRegistro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6090
   ClientLeft      =   810
   ClientTop       =   1755
   ClientWidth     =   8595
   Icon            =   "frmLogEnviosRegistro.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8595
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      TabIndex        =   31
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   5220
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   2820
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4154
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   5640
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7200
      TabIndex        =   26
      Top             =   5640
      Width           =   1275
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1440
      TabIndex        =   25
      Top             =   5640
      Width           =   1275
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remitente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1020
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   8355
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   285
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   5235
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
         Height          =   250
         Left            =   2640
         TabIndex        =   2
         Top             =   330
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   280
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1635
      End
      Begin VB.TextBox txtUbigeoOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   6855
      End
      Begin VB.TextBox txtUbiOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   630
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ubic.Geográfica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   23
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8355
      Begin VB.TextBox txtUbigeoDest 
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   5235
      End
      Begin VB.CommandButton cmdUbiDest 
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
         Height          =   250
         Left            =   2610
         TabIndex        =   18
         Top             =   930
         Width           =   315
      End
      Begin VB.TextBox txtDireccionDest 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   570
         Width           =   6855
      End
      Begin VB.TextBox txtPersonaDest 
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   5235
      End
      Begin VB.CommandButton cmdPersCod 
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
         Height          =   250
         Left            =   2610
         TabIndex        =   7
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtPersDest 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtUbiDest 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubic.Geográfica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   21
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destinatario"
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
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   120
      TabIndex        =   13
      Top             =   2175
      Width           =   8355
      Begin VB.TextBox txtPersonaEnvio 
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   5235
      End
      Begin VB.CommandButton cmdPersEnvio 
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
         Height          =   250
         Left            =   2610
         TabIndex        =   14
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtPersEnvio 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa envío"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5880
      TabIndex        =   30
      Top             =   5310
      Width           =   615
   End
End
Attribute VB_Name = "frmLogEnviosRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String
Dim sSQL As String, nEditable As Boolean
Dim cAgeCod As String, cAreaCod As String, cCargoCod As String

Public Sub Inicio(ByVal psOpeCod As String, Optional pnEditable As Boolean = False)
cOpeCod = psOpeCod
nEditable = pnEditable
Me.Show 1
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim nMovReg As Integer
Dim nFilas As Integer
Dim i As Integer
Dim c

nFilas = MSFlex.Rows - 1

If MsgBox("¿ Está seguro de grabar los datos mostrados ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then

   If Not oConn.AbreConexion Then
      MsgBox "No se puede establecer la conexión..." + Space(10), vbInformation, "Aviso"
      Exit Sub
   End If
   
   sSQL = "INSERT INTO LogCtrlEnvios (dFecha,cPersCod,cAreaCod,cAgeCod,cUbigeoOrigen,cPersDestino,cDireccionDestino,cUbigeoDestino,cPersEnvio ) " & _
          " VALUES ('" & Format(Date, "YYYYMMDD") & "','" & txtPersCod.Text & "','" & cAreaCod & "', '" & cAgeCod & "','" & txtUbiOrig.Text & "','" & txtPersDest.Text & "','" & txtDireccionDest.Text & "','" & txtUbiDest.Text & "','" & txtPersEnvio.Text & "') "
   oConn.Ejecutar sSQL
   
   nMovReg = UltimaSecuenciaIdentidad("LogCtrlEnvios")
   
   For i = 1 To nFilas
       sSQL = "INSERT INTO LogCtrlEnviosDetalle (nMovReg,nMovRegItem,nTipoEnvio,cComentario,nCostoEnvio)  " & _
              "  VALUES (" & nMovReg & ", " & CInt(Val(MSFlex.TextMatrix(i, 0))) & ", " & CInt(Val(MSFlex.TextMatrix(i, 1))) & ", '" & MSFlex.TextMatrix(i, 3) & "', " & VNumero(MSFlex.TextMatrix(i, 4)) & " )"
       oConn.Ejecutar sSQL
   Next
   
   oConn.CierraConexion
   
   MsgBox "Se han grabado los datos correctamente!" + Space(10), vbInformation, "Aviso"
   Unload Me
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
cmdPersona.Visible = True
txtPersEnvio.Text = "1120600442666"
txtPersonaEnvio.Text = "TRANSPORTES LINEA S.A."
CompruebaUsuario
FormaFlex
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 400:       MSFlex.TextMatrix(0, 0) = "Item": MSFlex.ColAlignment(0) = 4
MSFlex.ColWidth(1) = 0
MSFlex.ColWidth(2) = 1300:      MSFlex.TextMatrix(0, 2) = "Tipo Envío"
MSFlex.ColWidth(3) = 5000:      MSFlex.TextMatrix(0, 3) = "Descripción del contenido"
MSFlex.ColWidth(4) = 1400:      MSFlex.TextMatrix(0, 4) = "Costo Envío"
MSFlex.ColWidth(5) = 0:         MSFlex.TextMatrix(0, 5) = ""
MSFlex.ColWidth(6) = 0
MSFlex.ColWidth(7) = 0
End Sub

Sub CompruebaUsuario()
txtPersCod = gsCodPersUser
txtPersona.Text = gsNomUser
If gsCodArea = "" Or gsCodUser = "SIST" Then
   cmdPersona.Visible = True
Else
   cmdPersona.Visible = False
End If
End Sub

Private Sub cmdAgregar_Click()
Dim k As Integer, nUltFila As Integer
Dim nTarifa As Currency

If Len(txtUbiOrig.Text) <= 4 Or Len(txtUbiDest.Text) <= 4 Then
   MsgBox "Debe indicar el Origen y Destino de la correspondencia..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

nUltFila = MSFlex.Rows - 1

If Len(MSFlex.TextMatrix(nUltFila, 1)) > 0 Then
   k = nUltFila
Else
   FormaFlex
   k = 0
End If

sSQL = "Select nConsValor, cConsDescripcion from Constante where nConsCod = 9135 and nConsCod <> nConsValor"
frmLogSelector.Consulta sSQL, "Seleccione Tipo de Envío"
If frmLogSelector.vpHaySeleccion Then
   nTarifa = ObtenerTarifaEnvio(frmLogSelector.vpCodigo)
   If nTarifa <= 0 Then
      MsgBox "No se puede hallar una tarifa para el envío de: " & Space(10) & vbCrLf & txtUbigeoOrig.Text & " -> " & txtUbigeoDest.Text & "..." + Space(10) + vbCrLf + vbCrLf + "    Consulte con el encargado de Logística", vbInformation, "Aviso"
      Exit Sub
   End If
   k = k + 1
   InsRow MSFlex, k
   MSFlex.TextMatrix(k, 1) = frmLogSelector.vpCodigo
   MSFlex.TextMatrix(k, 2) = frmLogSelector.vpDescripcion
   MSFlex.TextMatrix(k, 4) = FNumero(nTarifa)
   For k = 1 To MSFlex.Rows - 1
       MSFlex.TextMatrix(k, 0) = k
   Next
   TotalFila 4
End If
End Sub

Function ObtenerTarifaEnvio(ByVal psCodigo As Integer) As Currency
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim nVeces As Integer

nVeces = 0
ObtenerTarifaEnvio = 0
If oConn.AbreConexion Then

   sSQL = "select nTarifaMonto from LogCtrlEnviosTarifas " & _
          " where nTipoEnvio=" & psCodigo & " and " & _
          "       '" & txtUbiOrig.Text & "' like SUBSTRING(cUbigeoOrigen,1,5)+'%' and " & _
          "       '" & txtUbiDest.Text & "' like SUBSTRING(cUbigeoDestino,1,5)+'%' "
          
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         nVeces = nVeces + 1
         ObtenerTarifaEnvio = rs!nTarifaMonto
         rs.MoveNext
      Loop
      If nVeces > 1 Then
         MsgBox "Existe mas de un costo para el envío...revise su Tarifario" + Space(10), vbInformation, "Aviso"
         ObtenerTarifaEnvio = 0
      End If
   End If
   oConn.CierraConexion
End If
End Function

Private Sub cmdPersCod_Click()
Dim x As UPersona
Set x = frmBuscaPersona.Inicio

If x Is Nothing Then
    Exit Sub
End If

If Len(Trim(x.sPersNombre)) > 0 Then
   txtPersonaDest.Text = x.sPersNombre
   txtPersDest.Text = x.sPersCod
End If
End Sub

Private Sub cmdPersEnvio_Click()
Dim x As UPersona
Set x = frmBuscaPersona.Inicio

If x Is Nothing Then
    Exit Sub
End If

If Len(Trim(x.sPersNombre)) > 0 Then
   txtPersonaEnvio.Text = x.sPersNombre
   txtPersEnvio = x.sPersCod
End If
End Sub

Private Sub cmdPersona_Click()
Dim x As UPersona
Set x = frmBuscaPersona.Inicio(True)

If x Is Nothing Then
    Exit Sub
End If

If Len(Trim(x.sPersNombre)) > 0 Then
   txtPersona.Text = x.sPersNombre
   txtPersCod = x.sPersCod
End If
End Sub

Private Sub txtPersCod_Change()
If Len(Trim(txtPersCod)) = 13 Then
   RecuperaDatosPersona
End If
End Sub

Sub RecuperaDatosPersona()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim cAnioMes As String

If Len(Trim(txtPersCod)) > 0 Then
   
   cAnioMes = CStr(Year(Date)) + Format(Month(Date), "00")
   sSQL = "select cRHCargoCodOficial,cRHAreaCodOficial,cRHAgenciaCodOficial from RHCargos where " & _
          "       cPersCod = '" & txtPersCod & "' and dRHCargoFecha in (select max(dRHCargoFecha) from RHCargos where " & _
          "       cPersCod  = '" & txtPersCod & "' and  dRHCargoFecha <= '" & cAnioMes & "')"
          
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSQL)
      If Not rs.EOF Then
         cAgeCod = rs!cRHAgenciaCodOficial
         cAreaCod = rs!cRHAreaCodOficial
         cCargoCod = rs!cRHCargoCodOficial
         
         Set rs = Nothing
         
         sSQL = "select a.cUbigeoCod, a.cAgeDescripcion, u.cUbigeoDescripcion " & _
         "  from Agencias a inner join UbicacionGeografica u on a.cUbigeoCod = u.cUbigeoCod " & _
         " where a.cAgeCod = '" & cAgeCod & "'"
 
         Set rs = oConn.CargaRecordSet(sSQL)
         If Not rs.EOF Then
            txtUbiOrig.Text = rs!cUbigeoCod
            txtUbigeoOrig.Text = rs!cAgeDescripcion + " - " + Trim(rs!cUbigeoDescripcion)
         End If
      End If
   End If
End If
End Sub

Private Sub cmdQuitar_Click()
Dim k As Integer
End Sub

Private Sub cmdUbiDest_Click()
frmLogProSelSeleUbiGeo.Show 1
    With frmLogProSelSeleUbiGeo
        txtUbiDest.Text = .gvCodigo
        txtUbigeoDest.Text = Trim(.gvNoddo)
    End With
End Sub

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************


Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
Dim k As Integer
If KeyCode = vbKeyDelete Then
   If MSFlex.Rows - 1 > 1 Then
      MSFlex.RemoveItem MSFlex.row
   Else
      MSFlex.RowHeight(1) = 8
      For k = 1 To MSFlex.Cols - 1
          MSFlex.TextMatrix(1, k) = ""
      Next
   End If
   For k = 1 To MSFlex.Rows - 1
       MSFlex.TextMatrix(k, 0) = k
   Next
   MSFlex.SetFocus
End If
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col = 3 And nEditable Then
   EditaFlex MSFlex, txtEdit, KeyAscii
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
'If InStr("0123456789", Chr(KeyAscii)) Then
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
Edt.Visible = True
Edt.SetFocus
'End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'nKeyAscii = KeyAscii
'KeyAscii = DigNumEnt(KeyAscii)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   'txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSFlex_GotFocus()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
'TotalFila MSFlex.row
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
'TotalFila MSFlex.row
End Sub

Sub TotalFila(i As Integer)
Dim j As Integer, n As Integer
Dim nSuma As Currency
nSuma = 0
For j = 1 To MSFlex.Rows - 1
    nSuma = nSuma + VNumero(MSFlex.TextMatrix(j, 4))
Next
txtTotal = FNumero(nSuma)
End Sub

