VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogEnviosAprobacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobación / Rechazo de Envíos "
   ClientHeight    =   6090
   ClientLeft      =   450
   ClientTop       =   1905
   ClientWidth     =   10935
   Icon            =   "frmLogEnviosAprobacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10935
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   7800
      TabIndex        =   9
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   330
         Left            =   495
         TabIndex        =   13
         Top             =   680
         Width           =   2415
      End
      Begin VB.TextBox txtFecFin 
         Height          =   300
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   11
         Top             =   330
         Width           =   1035
      End
      Begin VB.TextBox txtFecIni 
         Height          =   300
         Left            =   480
         MaxLength       =   10
         TabIndex        =   10
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Requerimientos "
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
         Left            =   840
         TabIndex        =   14
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del                           Al"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   390
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Procesar  Aprobaciones y/o Rechazos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   5640
      Width           =   3675
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9660
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
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
      Height          =   1080
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7635
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   7275
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
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Top             =   330
         Width           =   350
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   315
         Width           =   5655
      End
      Begin VB.TextBox txtPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Height          =   280
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Persona que realiza Aprobación / Rechazo "
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
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   -15
         Width           =   3735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle del Envío "
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
      Height          =   2295
      Left            =   60
      TabIndex        =   15
      Top             =   3240
      Width           =   10815
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtDestinatario 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   300
         Width           =   7095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDet 
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   630
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destinatario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   20
         Top             =   345
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8640
         TabIndex        =   19
         Top             =   375
         Width           =   390
      End
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmLogEnviosAprobacion.frx":08CA
      Top             =   5700
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmLogEnviosAprobacion.frx":0C0C
      Top             =   5700
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmLogEnviosAprobacion.frx":0F4E
      Top             =   5700
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuQuitar 
      Caption         =   "MenuObs"
      Visible         =   0   'False
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular Item"
      End
   End
End
Attribute VB_Name = "frmLogEnviosAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cOpeCod As String, cOpeDesc As String
Dim sSQL As String, PuedeMarcar As Boolean
Dim cAgeCod As String, cAreaCod As String, cCargoCod As String

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
cOpeCod = psOpeCod
cOpeDesc = psOpeDesc
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = cOpeDesc
cAreaCod = ""
cAgeCod = ""
PuedeMarcar = True
txtFecIni.Text = "01/" & Format(Month(Date), "00") & "/" + CStr(Year(Date))
txtFecFin.Text = Date
CompruebaUsuario
FormaFlex
FormaFlexDet
End Sub

Private Sub cmdSalir_Click()
Unload Me
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

Private Sub cmdBuscar_Click()
If Len(cAreaCod) = 0 Or Len(cAgeCod) = 0 Then
   MsgBox "La persona no es válida para realizar esta operación..." + Space(10), vbInformation, "Aviso"
Else
   ListaEnviosArea cAreaCod, cAgeCod
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

Sub RecuperaDatosPersona()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, cAnioMes As String

cAgeCod = ""
cAreaCod = ""
cCargoCod = ""

If Len(Trim(txtPersCod)) > 0 Then
   cAnioMes = CStr(Year(Date)) + Format(Month(Date), "00")
   
   sSQL = "select r.cRHCargoCodOficial,r.cRHAreaCodOficial,r.cRHAgenciaCodOficial,cArea = coalesce(a.cAreaDescripcion,'')  " & _
          "  from RHCargos r left join Areas a on r.cRHAreaCodOficial = a.cAreaCod where " & _
          "       r.cPersCod = '" & txtPersCod & "' and r.dRHCargoFecha in (select max(dRHCargoFecha) from RHCargos where " & _
          "       cPersCod  = '" & txtPersCod & "' and  dRHCargoFecha <= '" & cAnioMes & "') "
          
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSQL)
      If Not rs.EOF Then
         txtArea.Text = rs!cArea
         cAgeCod = rs!cRHAgenciaCodOficial
         cAreaCod = rs!cRHAreaCodOficial
         cCargoCod = rs!cRHCargoCodOficial
      End If
   End If
End If
End Sub

Sub ListaEnviosArea(psAreaCod As String, psAgeCod As String)
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim cConsulta As String

i = 0
FormaFlex
FormaFlexDet

If cOpeCod = "541032" Then

   cConsulta = " and e.cAreaCod in (SELECT cAreaCod FROM Areas a inner join (SELECT cAreaEstruc FROM Areas where cAreaCod = '" & psAreaCod & "') b on a.cAreaEstruc like b.cAreaEstruc+'%') and e.cAgeCod = '" & psAgeCod & "'  and e.nEstado = 1 and len(e.cPersCodJefe)=0 "
   
ElseIf cOpeCod = "541033" Then

   cConsulta = " and e.nEstado = 2 and len(e.cPersCodJefe)>0 and len(e.cPersCodLog)=0 "
   
   If psAreaCod <> "036" Then
      MsgBox "Sólo personal autorizado del área de Logística " + Space(10) + vbCrLf + _
             "        puede realizar esta operación...", vbInformation, "Aviso"
      PuedeMarcar = False
      cmdGrabar.Enabled = False
      Exit Sub
   Else
      PuedeMarcar = True
      cmdGrabar.Enabled = True
   End If
   
End If

sSQL = "select e.nMovReg, e.dFecha,p.cPersNombre as cPersona, d.cPersNombre as cPersDestino, e.cDireccionDestino as cDireccion, n.nCosto ," & _
       "       uo.cUbigeoDescripcion as cOrigen, ud.cUbigeoDescripcion as cDestino" & _
       "  from LogCtrlEnvios e inner join Persona p on e.cPersCod = p.cPersCod " & _
       "                       inner join Persona d on e.cPersDestino = d.cPersCod " & _
       "       inner join UbicacionGeografica uo on e.cUbigeoOrigen = uo.cUbigeoCod " & _
       "       inner join UbicacionGeografica ud on e.cUbigeoDestino = ud.cUbigeoCod inner join (select nMovReg, sum(nCostoEnvio) as nCosto from LogCtrlEnviosDetalle group by nMovReg) n on e.nMovReg = n.nMovReg " & _
       " WHERE e.dFecha >= '" & Format(txtFecIni.Text, "YYYYMMDD") & "' and e.dFecha <= '" & Format(txtFecFin.Text, "YYYYMMDD") & "' " & _
       "    " & cConsulta & _
       "    "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 2) = rs!nMovReg
         MSFlex.TextMatrix(i, 3) = rs!cPersona
         MSFlex.TextMatrix(i, 4) = rs!cOrigen
         MSFlex.TextMatrix(i, 5) = rs!cPersDestino
         MSFlex.TextMatrix(i, 6) = rs!cDestino
         MSFlex.TextMatrix(i, 7) = FNumero(rs!nCosto)
         MSFlex.row = i
         MSFlex.Col = 0
         Set MSFlex.CellPicture = imgCheck(0).Picture
         rs.MoveNext
      Loop
   End If
End If

sSQL = "select cRHCargoCod,cRHCargoDescripcion,nRHFuncionario from RHCargosTabla " & _
       " where nRHFuncionario in (1,2,4) "

sSQL = "select cRHCargoCod,cRHCargoDescripcion,nRHFuncionario from RHCargosTabla " & _
       " where nRHFuncionario not in (1,2,4)"

End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 260
MSFlex.ColWidth(1) = 0
MSFlex.ColWidth(2) = 0
MSFlex.ColWidth(3) = 3400:      MSFlex.TextMatrix(0, 3) = "Remitente"
MSFlex.ColWidth(4) = 2800:      MSFlex.TextMatrix(0, 4) = ""
MSFlex.ColWidth(5) = 0  ':      MSFlex.TextMatrix(0, 5) = "Destinatario"
MSFlex.ColWidth(6) = 2800:      MSFlex.TextMatrix(0, 6) = ""
MSFlex.ColWidth(7) = 1200
End Sub

Private Sub mnuAnular_Click()
Dim k As Integer, oConn As New DConecta
Dim nMov As Integer, nItem As Integer
k = MSDet.row

If CInt(Val(MSDet.TextMatrix(MSDet.row, 0))) <= 0 Or CInt(Val(MSDet.TextMatrix(MSDet.row, 1))) <= 0 Then Exit Sub

nMov = MSDet.TextMatrix(MSDet.row, 0)
nItem = MSDet.TextMatrix(MSDet.row, 1)


sSQL = "UPDATE LogCtrlEnviosDetalle SET nEstadoItem = 0  " & _
       " WHERE nMovReg = " & nMov & " and nMovRegItem = " & nItem & " " & _
       "" & _
       ""
       
If MsgBox("¿ Está seguro de anular el Item indicado ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
      ListaEnvioDetalle nMov
   End If
End If
End Sub

Private Sub MSDet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Len(MSDet.TextMatrix(MSDet.row, 0)) > 0 Then
   PopupMenu mnuQuitar
End If
End Sub

Private Sub MSFlex_DblClick()
Dim k As Integer, n As Integer
If PuedeMarcar Then
   k = MSFlex.row
   n = Len(MSFlex.TextMatrix(k, 1))
   MSFlex.row = k
   MSFlex.Col = 0
   Select Case n
    Case 0
         MSFlex.TextMatrix(k, 1) = "."
         Set MSFlex.CellPicture = imgCheck(1).Picture
    Case 1
         MSFlex.TextMatrix(k, 1) = ".."
         Set MSFlex.CellPicture = imgCheck(2).Picture
    Case 2
         MSFlex.TextMatrix(k, 1) = ""
         Set MSFlex.CellPicture = imgCheck(0).Picture
   End Select
   MSFlex.Col = 3
End If
End Sub

Private Sub MSFlex_GotFocus()
If CInt(Val(MSFlex.TextMatrix(MSFlex.row, 2))) > 0 Then
   ListaEnvioDetalle MSFlex.TextMatrix(MSFlex.row, 2)
End If
End Sub

Private Sub MSFlex_RowColChange()
If CInt(Val(MSFlex.TextMatrix(MSFlex.row, 2))) > 0 Then
   ListaEnvioDetalle MSFlex.TextMatrix(MSFlex.row, 2)
End If
End Sub

Sub ListaEnvioDetalle(pnMovReg As Integer)
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim nSuma As Currency

i = 0
nSuma = 0
FormaFlexDet

sSQL = "select d.nMovRegItem, t.cConsDescripcion, d.cComentario, d.nCostoEnvio, d.nEstadoItem " & _
       "  from LogCtrlEnviosDetalle d " & _
       " inner join (Select  nConsValor, cConsDescripcion from Constante where nConsCod =  9135 and nConsCod<>nConsValor) t on d.nTipoEnvio = t.nConsValor " & _
       " where d.nMovReg = " & pnMovReg & " and d.nEstadoItem = 1 order by d.nMovRegItem " & _
       "" & _
       ""

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSDet, i
         MSDet.TextMatrix(i, 0) = pnMovReg
         MSDet.TextMatrix(i, 1) = rs!nMovRegItem
         MSDet.TextMatrix(i, 2) = rs!cConsDescripcion
         MSDet.TextMatrix(i, 3) = rs!cComentario
         MSDet.TextMatrix(i, 4) = FNumero(rs!nCostoEnvio)
         MSDet.TextMatrix(i, 5) = rs!nEstadoItem
         'nSuma = nSuma + rs!nCostoEnvio
         rs.MoveNext
      Loop
      'txtTotal.Text = FNumero(nSuma)
      txtDestinatario.Text = MSFlex.TextMatrix(MSFlex.row, 5)
      txtTotal.Text = MSFlex.TextMatrix(MSFlex.row, 7)
   End If
End If
End Sub

Sub FormaFlexDet()
MSDet.Clear
MSDet.Rows = 2
MSDet.RowHeight(0) = 320
MSDet.RowHeight(1) = 8
MSDet.ColWidth(0) = 0
MSDet.ColWidth(1) = 400:       MSDet.TextMatrix(0, 1) = "Item":        MSDet.ColAlignment(1) = 4
MSDet.ColWidth(2) = 2350:      MSDet.TextMatrix(0, 2) = "Tipo Envio"
MSDet.ColWidth(3) = 6000:      MSDet.TextMatrix(0, 3) = "Descripción"
MSDet.ColWidth(4) = 1500:      MSDet.TextMatrix(0, 4) = " Costo envío"
MSDet.ColWidth(5) = 0
MSDet.ColWidth(6) = 0
MSDet.ColWidth(7) = 0
txtDestinatario.Text = ""
txtTotal.Text = ""
End Sub


Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim i As Integer, k As Integer, n As Integer
Dim nEstado As Integer, nMovReg As Integer

n = MSFlex.Rows - 1

If MsgBox("¿ Seguro de grabar los envíos de correspondencia ?" + Space(10), vbYesNo + vbQuestion, "Confirme") = vbYes Then

   If Not oConn.AbreConexion Then
      Exit Sub
   End If
   
   If cOpeCod = "541032" Then
      For i = 1 To n
          If Len(MSFlex.TextMatrix(i, 1)) > 0 Then
             nEstado = Len(MSFlex.TextMatrix(i, 1)) + 1
             nMovReg = CInt(Val(MSFlex.TextMatrix(i, 2)))
             sSQL = "UPDATE LogCtrlEnvios SET nEstado = " & nEstado & ", cPersCodJefe = '" & txtPersCod.Text & "' WHERE nMovReg = " & nMovReg & " "
             oConn.Ejecutar sSQL
          End If
      Next
   
   ElseIf cOpeCod = "541033" Then
   
      For i = 1 To n
          If Len(MSFlex.TextMatrix(i, 1)) > 0 Then
             nMovReg = CInt(Val(MSFlex.TextMatrix(i, 2)))
             sSQL = "UPDATE LogCtrlEnvios SET nEstado = 4, " & _
                    "       dFechaEnvio = '" & Format(Date, "YYYYMMDD") & "', " & _
                    "       cPersCodLog = '" & txtPersCod.Text & "' " & _
                    " WHERE nMovReg = " & nMovReg & " "
             oConn.Ejecutar sSQL
          End If
      Next
   
   End If
   
   ListaEnviosArea cAreaCod, cAgeCod
   oConn.CierraConexion
End If
End Sub

Private Sub txtFecIni_GotFocus()
SelTexto txtFecIni
End Sub

Private Sub txtFecFin_GotFocus()
SelTexto txtFecFin
End Sub

Private Sub txtPersCod_Change()
If Len(Trim(txtPersCod)) = 13 Then
   RecuperaDatosPersona
End If
End Sub
