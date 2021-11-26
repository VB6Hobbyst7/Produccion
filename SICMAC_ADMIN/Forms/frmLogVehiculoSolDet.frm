VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoSolDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Vehicular - Detalle de Solicitud de Asignación Vehicular "
   ClientHeight    =   4290
   ClientLeft      =   1335
   ClientTop       =   1680
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8805
   Begin VB.Frame Frame3 
      Caption         =   "Kilometraje "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   30
      Top             =   2640
      Width           =   1635
      Begin VB.TextBox txtKmIni 
         Height          =   315
         Left            =   180
         MaxLength       =   15
         TabIndex        =   31
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Lectura Actual"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   23
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdVehiculo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdConductor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   26
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtVehiculo 
         Appearance      =   0  'Flat
         Height          =   280
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   6795
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   6795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vehículo"
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
         Left            =   180
         TabIndex        =   29
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conductor"
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
         TabIndex        =   28
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdAsigna 
      Caption         =   "Grabar Solicitud de Asignación"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Asignación "
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
      Height          =   1095
      Left            =   60
      TabIndex        =   18
      Top             =   2640
      Width           =   1995
      Begin VB.OptionButton opAsig1 
         Caption         =   " Indefinida"
         Height          =   255
         Left            =   300
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton opAsig2 
         Caption         =   " Temporal"
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   4815
      Begin VB.TextBox txtHoraIni 
         Height          =   315
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   15
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtMinIni 
         Height          =   315
         Left            =   3420
         TabIndex        =   14
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtFechaIni 
         Height          =   315
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   13
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox cboHoraIni 
         Height          =   315
         ItemData        =   "frmLogVehiculoSolDet.frx":0000
         Left            =   3840
         List            =   "frmLogVehiculoSolDet.frx":000A
         TabIndex        =   12
         Top             =   300
         Width           =   735
      End
      Begin VB.Frame fraFin 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox txtFechaFin 
            Height          =   315
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   9
            Top             =   0
            Width           =   1155
         End
         Begin VB.TextBox txtHoraFin 
            Height          =   315
            Left            =   2820
            TabIndex        =   8
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtMinFin 
            Height          =   315
            Left            =   3300
            TabIndex        =   7
            Top             =   0
            Width           =   375
         End
         Begin VB.ComboBox cboHoraFin 
            Height          =   315
            ItemData        =   "frmLogVehiculoSolDet.frx":0016
            Left            =   3720
            List            =   "frmLogVehiculoSolDet.frx":0020
            TabIndex        =   6
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta  el dia"
            Height          =   195
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Hora            :"
            Height          =   195
            Left            =   2340
            TabIndex        =   10
            Top             =   60
            Width           =   930
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hora            :"
         Height          =   195
         Left            =   2460
         TabIndex        =   17
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde el dia"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Responsables "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   8655
      Begin VB.CommandButton cmdQuitarR 
         Caption         =   "Quitar"
         Height          =   350
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregaR 
         Caption         =   "Agregar"
         Height          =   350
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   855
         Left            =   1620
         TabIndex        =   3
         Top             =   300
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1508
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
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
         _Band(0).Cols   =   3
      End
   End
End
Attribute VB_Name = "frmLogVehiculoSolDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpGrabado As Boolean
Dim cPersCod As String, nVehiculoCod As Integer
Dim cAgeCod As String
Dim nKeyAscii As Integer

Public Sub Agencia(vAgeCod As String)
cAgeCod = vAgeCod
Me.Show 1
End Sub

Private Sub cmdAgregar_Click()
Dim k As Integer, X As UPersona

Set X = frmBuscaPersona.Inicio

k = MSFlex.Rows - 1
If Len(Trim(MSFlex.TextMatrix(k, 1))) = 0 Then
   k = 0
End If

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   'txtPersona.Text = X.sPersNombre
   'txtPersCod = X.sPersCod
   k = k + 1
   InsRow MSFlex, k
   MSFlex.TextMatrix(k, 1) = X.sPersCod
   MSFlex.TextMatrix(k, 2) = X.sPersNombre
End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   k = k + 1
'   InsRow MSFlex, k
'   MSFlex.TextMatrix(k, 1) = frmBuscaPersona.vpPersCod
'   MSFlex.TextMatrix(k, 2) = frmBuscaPersona.vpPersNom
'End If
opAsig1.SetFocus
End Sub

Private Sub cmdQuitarR_Click()
Dim k As Integer
k = MSFlex.row

If Len(Trim(MSFlex.TextMatrix(k, 1))) = 0 Then Exit Sub


If MsgBox("¿Seguro de quitar la persona ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If k > 1 Then
      MSFlex.RemoveItem k
   Else
      MSFlex.TextMatrix(1, 1) = ""
      MSFlex.TextMatrix(1, 2) = ""
      MSFlex.RowHeight(1) = 8
   End If
End If
End Sub

Private Sub Form_Load()
CentraForm Me
FlexResp
Me.vpGrabado = False
txtFechaIni = Date
txtFechaFin = Date
cboHoraIni.ListIndex = 0
cboHoraFin.ListIndex = 1
txtHoraIni.Text = "08":    txtMinIni.Text = "00"
txtHoraFin.Text = "08":    txtMinFin.Text = "00"
End Sub

Private Sub cmdConductor_Click()
Dim sSQL As String

sSQL = "select v.cPersCod,cPersona=replace(p.cPersNombre,'/',' ') " & _
       "  from LogVehiculoConductor v inner join Persona p on p.cPersCod = v.cPersCod  " & _
       " where v.cAgeCod = '" & cAgeCod & "' and v.nEstado=1"
       
frmLogSelector.Consulta sSQL, "Seleccione Conductor"
If frmLogSelector.vpHaySeleccion Then
   cPersCod = frmLogSelector.vpCodigo
   txtPersona = frmLogSelector.vpDescripcion
   cmdVehiculo.SetFocus
End If
End Sub

Private Sub CmdSalir_Click()
Me.vpGrabado = False
Unload Me
End Sub

Private Sub cmdVehiculo_Click()
Dim sSQL As String

'Selecciona vehiculos con Estado = 1 : LIBRES
sSQL = "select v.nVehiculoCod,t.cDescripcion,v.cModelo,v.cPlaca " & _
       "  from LogVehiculo v " & _
       "  inner join (select nConsValor as nTipoVehiculo, cConsDescripcion as cDescripcion " & _
       "                from Constante where nConsCod = 9026 and nConsCod<>nConsValor) t on v.nTipoVehiculo = t.nTipoVehiculo " & _
       " where v.nEstado=1"
       
frmLogSelector.Consulta sSQL, "Seleccione Vehiculo"
If frmLogSelector.vpHaySeleccion Then
   nVehiculoCod = frmLogSelector.vpCodigo
   txtVehiculo = frmLogSelector.vpDescripcion
   cmdAgregar.SetFocus
   'txtFechaIni.SetFocus
End If
End Sub

Sub FlexResp()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 300
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 1100:  MSFlex.TextMatrix(0, 1) = "Código"
MSFlex.ColWidth(2) = 5000:  MSFlex.TextMatrix(0, 2) = "Responsable"
End Sub

Private Sub opAsig1_Click()
If opAsig1.value Then
   fraFin.Visible = False
Else
   fraFin.Visible = True
End If
End Sub

Private Sub opAsig2_Click()
If opAsig1.value Then
   fraFin.Visible = False
Else
   fraFin.Visible = True
End If
End Sub

Private Sub opAsig1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaIni.SetFocus
End If
End Sub

Private Sub opAsig2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaIni.SetFocus
End If
End Sub


'*****************************************************************

Private Sub txtFechaIni_GotFocus()
SelTexto txtFechaIni
End Sub

Private Sub txtFechaFin_GotFocus()
SelTexto txtFechaFin
End Sub

Private Sub txtHoraIni_GotFocus()
SelTexto txtHoraIni
End Sub

Private Sub txtHoraFin_GotFocus()
SelTexto txtHoraFin
End Sub

Private Sub txtMinIni_GotFocus()
SelTexto txtMinIni
End Sub

Private Sub txtMinFin_GotFocus()
SelTexto txtMinFin
End Sub

'*****************************************************************

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaIni, KeyAscii)
If nKeyAscii = 13 Then
   txtHoraIni.SetFocus
End If
End Sub

Private Sub txtHoraIni_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtHoraIni = Format(txtHoraIni, "00")
   txtMinIni.SetFocus
End If
End Sub

Private Sub txtMinIni_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtMinIni = Format(txtMinIni, "00")
   cboHoraIni.SetFocus
End If
End Sub

Private Sub cboHoraIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If fraFin.Visible Then
      txtFechaFin.SetFocus
   Else
      txtKmIni.SetFocus
   End If
End If
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaFin, KeyAscii)
If nKeyAscii = 13 Then
   txtHoraFin.SetFocus
End If
End Sub

Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtHoraFin = Format(txtHoraFin, "00")
   txtMinFin.SetFocus
End If
End Sub

Private Sub txtMinFin_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtMinFin = Format(txtMinFin, "00")
   cboHoraFin.SetFocus
End If
End Sub

Private Sub cboHoraFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtKmIni.SetFocus
End If
End Sub

Private Sub txtKmIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAsigna.SetFocus
End If
End Sub

'Private Sub txtKmFin_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   cmdAsigna.SetFocus
'End If
'End Sub

'*****************************************************************

Private Sub cmdAsigna_Click()
Dim oConn As New DConecta, i As Integer, n As Integer
Dim sSQL As String, nVCod As Integer, cCCod As String
Dim sMovNro As String, nEstado As Integer, sPersCod As String
Dim nNro As Integer, nAsignaNro As Integer, nTipo As Integer

nVCod = nVehiculoCod
cCCod = cPersCod

If opAsig1.value Then nTipo = 1
If opAsig2.value Then nTipo = 2

If nVCod = 0 Then
   MsgBox "No se indica el Vehiculo..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If Len(Trim(cCCod)) = 0 And txtPersona.Visible Then
   MsgBox "Debe indicar el Conductor..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

nNro = 0
n = MSFlex.Rows - 1
For i = 1 To n
    If Len(Trim(MSFlex.TextMatrix(i, 1))) > 0 Then
       nNro = nNro + 1
       Exit For
    End If
Next

If nNro = 0 Then
   MsgBox "Debe indicar por lo menos un Responsable..." + Space(10), vbInformation
   Exit Sub
End If

If Not IsDate(txtFechaIni) Then
   MsgBox "La fecha de inicio de asignación no es válida..." + Space(10), vbInformation
   Exit Sub
End If

If Not IsDate(txtFechaFin) Then
   MsgBox "La fecha Final de asignación no es válida..." + Space(10), vbInformation
   Exit Sub
End If

'nEstado = cboEstado.ItemData(cboEstado.ListIndex)
If MsgBox("¿ Está seguro de asignar el vehículo ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   sMovNro = GetLogMovNro
   nEstado = 2
   
   If Not oConn.AbreConexion Then
      MsgBox "No se puede establecer conexión..." + Space(10), vbInformation, "Aviso"
      Exit Sub
   End If
   
   If nTipo = 1 Then
      sSQL = "INSERT INTO LogVehiculoAsignacion (nVehiculoCod,nAsignacionTpo,cPersCod,dFechaIni,KmInicial,nEstado) " & _
             "       VALUES (" & nVCod & "," & nTipo & ",'" & cCCod & "','" & Format(txtFechaIni, "YYYYMMDD") & "'," & VNumero(txtKmIni.Text) & "," & nEstado & ") "
   End If
   
   If nTipo = 2 Then
      sSQL = "INSERT INTO LogVehiculoAsignacion (nVehiculoCod,nAsignacionTpo,cPersCod,dFechaIni,dFechaFin,KmInicial,nEstado) " & _
             "       VALUES (" & nVCod & "," & nTipo & ",'" & cCCod & "','" & Format(txtFechaIni, "YYYYMMDD") & "','" & Format(txtFechaFin, "YYYYMMDD") & "'," & VNumero(txtKmIni.Text) & "," & nEstado & ") "
   End If
   
   oConn.Ejecutar sSQL
   
   nAsignaNro = UltimaSecuenciaIdentidad("LogVehiculoAsignacion")
   
   sSQL = "INSERT INTO LogVehiculoAsignacionMov (nAsignacionNro,cOpeCod,cMovNro,cComentario) " & _
          "       VALUES (" & nAsignaNro & ", '" & gsOpeCod & "','" & sMovNro & "','') "
   oConn.Ejecutar sSQL

   n = MSFlex.Rows - 1
   For i = 1 To n
       sPersCod = MSFlex.TextMatrix(i, 1)
       If Len(Trim(sPersCod)) > 0 Then
          sSQL = "INSERT INTO LogVehiculoResponsable (nAsignacionNro,cPersCod) " & _
                 "       VALUES (" & nAsignaNro & ", '" & sPersCod & "') "
          oConn.Ejecutar sSQL
       End If
   Next

   sSQL = "UPDATE LogVehiculo SET nEstado=" & nEstado & " WHERE nVehiculoCod = " & nVCod & " "
   oConn.Ejecutar sSQL
   
   If Len(Trim(cCCod)) > 0 And txtPersona.Visible Then
      sSQL = "UPDATE LogVehiculoConductor SET nEstado=" & nEstado & " WHERE cPersCod = '" & cCCod & "' "
      oConn.Ejecutar sSQL
   End If
   
   Me.vpGrabado = True
   Unload Me
End If
End Sub

'*****************************************************************


'*****************************************************************



