VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLogVehiculoRegistro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   1320
   ClientTop       =   2160
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8895
   Begin VB.Frame Frame3 
      Caption         =   "Usuario "
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
      Height          =   1515
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   8655
      Begin VB.TextBox txtFechaIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtFechaFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtEstado 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1020
         Width           =   3255
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   7335
      End
      Begin VB.TextBox txtVehiculo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   660
         Width           =   7335
      End
      Begin VB.TextBox txtVehiculoCod 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   660
         Width           =   1440
      End
      Begin VB.TextBox txtPersCod 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Left            =   6660
         TabIndex        =   13
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vehículo"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conductor"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7500
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame fraVista 
      Height          =   3600
      Left            =   120
      TabIndex        =   14
      Top             =   1500
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7260
         TabIndex        =   17
         Top             =   3100
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdAcepta 
         Caption         =   "Acepto"
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         Top             =   3100
         Visible         =   0   'False
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox rtfVista 
         Height          =   2475
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4366
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmLogVehiculoRegistro.frx":0000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Declaro conocer y aceptar las condiciones del presente Reglamento"
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
         TabIndex        =   18
         Top             =   300
         Width           =   5835
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Incidencias"
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
      Height          =   3450
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   3000
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7260
         TabIndex        =   20
         Top             =   3000
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2655
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   11
      End
   End
End
Attribute VB_Name = "frmLogVehiculoRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAsignaNro As Integer, nEstado As Integer, sPersCod As String, bVisualiza As Boolean

Public Sub Estado(ByVal psPersCod As String, Optional pnEstado As Integer, Optional pbSoloVisualiza As Boolean = False)
nEstado = pnEstado
sPersCod = psPersCod
bVisualiza = pbSoloVisualiza
Me.Show 1
End Sub

Private Sub Form_Load()
Me.Caption = "Control Vehicular - Aceptación de asignación vehicular"
LimpiaFlex
If Len(Trim(sPersCod)) > 0 Then
   DeterminaForma sPersCod
   fraReg.Visible = True
   fraVista.Visible = False
Else
   MsgBox "No se puede hallar la persona de código [" + sPersCod + "]..." + Space(10), vbInformation
   cmdAcepta.Visible = False
   cmdCancelar.Visible = True
   fraReg.Visible = False
   fraVista.Visible = True
   'Me.Height = 2500
   CentraForm Me
End If
End Sub

Sub DeterminaForma(ByVal psPersCod As String)
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim sSQL As String

nAsignaNro = 0
txtPersCod.Text = ""
txtPersona.Text = ""
txtVehiculo.Text = ""
txtFechaIni.Text = ""
txtFechaIni.Text = ""
txtVehiculoCod.Text = ""

sSQL = "select a.nAsignacionNro, a.cPersCod, a.nVehiculoCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
       "       a.dFechaIni, a.dFechaFin,t.cDescripcion,r.cPlaca, a.nEstado " & _
       "  from LogVehiculoAsignacion a inner join Persona p on a.cPersCod = p.cPersCod " & _
       "  inner join LogVehiculo r on a.nVehiculoCod = r.nVehiculoCod " & _
       "  left join (select nConsValor as nTipoVehiculo, cConsDescripcion as cDescripcion from Constante where nConsCod = 9026 and nConsCod<>nConsValor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
       "  where a.cPersCod = '" & psPersCod & "' and a.nEstado = " & nEstado & " " & _
       " "
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not rs.EOF Then
      nAsignaNro = rs!nAsignacionNro
      txtPersCod.Text = rs!cPersCod
      txtPersona.Text = rs!cPersona
      txtVehiculoCod.Text = rs!nVehiculoCod
      txtVehiculo.Text = rs!cDescripcion + " - " + rs!cPlaca
      txtFechaIni.Text = rs!dFechaIni
      txtFechaFin.Text = rs!dFechaFin
      If rs!nEstado = 3 Then
         Me.Caption = "Control Vehicular - Aceptación de asignación vehicular"
         txtEstado.ForeColor = "&H00000080"
         txtEstado.Text = "ASIGNACION APROBADA"
         cmdAcepta.Visible = True
         cmdCancelar.Visible = True
         fraReg.Visible = False
         Me.Height = 2500
      End If
      If rs!nEstado = 4 Then
         Me.Caption = "Control Vehicular - Registro de Incidencias"
         txtEstado.ForeColor = "&H00C00000"
         txtEstado.Text = "ASIGNACION ACEPTADA"
         cmdAcepta.Visible = False
         cmdCancelar.Visible = False
         fraReg.Visible = True
         Me.Height = 5700
         DoEvents
         RecuperaDatos nAsignaNro
      End If
   Else
      cmdCancelar.Visible = True
      fraReg.Visible = False
      Me.Height = 2500
      
      If oConn.AbreConexion Then
         Set rs = oConn.CargaRecordSet("Select cPersNombre from Persona where cPersCod = '" & psPersCod & "'")
         If Not rs.EOF Then
            txtPersona.Text = rs!cPersNombre
         End If
      End If
      
      If nEstado = 3 Then
         MsgBox "No hay asignaciones aprobadas para ACEPTAR..." + Space(10), vbInformation, "Aviso"
      End If
      
      If nEstado = 4 Then
         MsgBox "No hay asignaciones Aceptadas para registrar incidencias..." + Space(10), vbInformation, "Aviso"
      End If
      
   End If
End If
If bVisualiza Then
   cmdAgregar.Visible = False
   cmdQuitar.Visible = False
End If
CentraForm Me
End Sub


'*********************************************************************************
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

'*********************************************************************************
Private Sub cmdAcepta_Click()
Dim oConn As New DConecta, sSQL1 As String, sSQL2 As String
Dim sMovNro As String

If MsgBox("¿ Está seguro de aceptar la Asignación ? " + Space(10), vbYesNo + vbQuestion, "Confirme operación") = vbYes Then

   sMovNro = GetLogMovNro
   
   'cPersCod = '" & txtPersCod.Text & "' and nVehiculoCod = " & CInt(txtVehiculoCod.Text) & "
   
   sSQL1 = "UPDATE LogVehiculoAsignacion SET nEstado = " & gcAceptado & " " & _
          " WHERE  nAsignacionNro = " & nAsignaNro & " " & _
          "  "
      
   sSQL2 = "insert into LogVehiculoAsignacionMov (nAsignacionNro,cOpeCod,cMovNro) " & _
          " values (" & nAsignaNro & ",'" & gsOpeCod & "','" & sMovNro & "' )"
   
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL1
      oConn.Ejecutar sSQL2
   End If

   MsgBox "La Aceptación de la asignación se realizó con éxito " + Space(10) + vbCrLf + " ahora puede registrar incidencias" + Space(10), vbInformation, "Aviso"
   Unload Me
End If
End Sub

Sub LimpiaFlex()
With Me.MSFlex
     .Rows = 2:              .RowHeight(0) = 300: .Clear
     .RowHeight(1) = 8
     .ColWidth(0) = 780:      .ColAlignment(0) = 4
     .ColWidth(1) = 0
     .ColWidth(2) = 0
     .ColWidth(3) = 0
     .ColWidth(4) = 0
     .ColWidth(5) = 2100:    .TextMatrix(0, 5) = "Operación"
     .ColWidth(6) = 2000:    .TextMatrix(0, 6) = "Descripción"
     .ColWidth(7) = 1150:    .ColAlignment(7) = 1
     .ColWidth(8) = 1150:    .ColAlignment(8) = 1
     .ColWidth(9) = 900:
     .ColWidth(10) = 0:
End With
End Sub

Private Sub cmdAgregar_Click()
Dim oConn As New DConecta
Dim k As Integer, nTipoReg As Integer, cDescrip As String
Dim cValor0 As String, cValor1 As String, cValor2 As String
Dim nMonto As Currency, sSQL As String
Dim dFecha As String, sMovNro As String

On Error GoTo Sal_Agregar

If nAsignaNro <= 0 Then
   MsgBox "No se halla un movimiento de registro vehicular..." + Space(10), vbInformation
   Exit Sub
End If

k = MSFlex.Rows - 1
If Len(MSFlex.TextMatrix(k, 1)) = 0 And Len(MSFlex.TextMatrix(k, 2)) = 0 Then
   k = k - 1
End If
frmLogVehiculoDatos.vpFecha = txtFechaIni.Text
frmLogVehiculoDatos.Show 1
If frmLogVehiculoDatos.vpAcepta Then
   dFecha = frmLogVehiculoDatos.vpFecha
   nTipoReg = frmLogVehiculoDatos.vpTipoReg
   cValor0 = frmLogVehiculoDatos.vpCodigo0
   cValor1 = frmLogVehiculoDatos.vpCodigo1
   cValor2 = frmLogVehiculoDatos.vpCodigo2
   cDescrip = frmLogVehiculoDatos.vpDescrip
   nMonto = frmLogVehiculoDatos.vpMonto
   
   If Not IsDate(dFecha) Then
      MsgBox "La fecha del movimiento es no válida..." + Space(10), vbInformation
      Exit Sub
   End If
   If MsgBox("¿ Está seguro de agregar el registro ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then
   
      HayTrans = False
      If Not oConn.AbreConexion Then
         MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
         Exit Sub
      End If
      
      oConn.BeginTrans
      HayTrans = True
      
      sMovNro = GetLogMovNro
      
      sSQL = "INSERT INTO LogVehiculoAsignacionDet (nAsignacionNro,dFecha,nTipoReg,cDescripcion,cValor0,cValor1,cValor2,nMonto,nEstado,cMovNro) " & _
             " VALUES (" & nAsignaNro & ",'" & Format(dFecha, "YYYYMMDD") & "'," & nTipoReg & ",'" & Mid(cDescrip, 1, 40) & "','" & cValor0 & "','" & cValor1 & "','" & cValor2 & "'," & nMonto & ",1,'" & sMovNro & "')"
      oConn.Ejecutar sSQL
      
      oConn.CommitTrans
      oConn.CierraConexion

      k = k + 1
      InsRow MSFlex, k
      '---- LO QUE SE GRABA --------------------------------------
      MSFlex.TextMatrix(k, 0) = frmLogVehiculoDatos.vpFecha
      MSFlex.TextMatrix(k, 1) = frmLogVehiculoDatos.vpTipoReg
      MSFlex.TextMatrix(k, 2) = frmLogVehiculoDatos.vpCodigo0
      MSFlex.TextMatrix(k, 3) = frmLogVehiculoDatos.vpCodigo1
      MSFlex.TextMatrix(k, 4) = frmLogVehiculoDatos.vpCodigo2
      MSFlex.TextMatrix(k, 6) = frmLogVehiculoDatos.vpDescrip
      MSFlex.TextMatrix(k, 9) = frmLogVehiculoDatos.vpMonto
    
      '----- LO QUE SE VISUALIZA ---------------------------------
      MSFlex.TextMatrix(k, 5) = frmLogVehiculoDatos.vpRegistro
      If frmLogVehiculoDatos.vpTipoReg = 4 Then
         MSFlex.TextMatrix(k, 7) = frmLogVehiculoDatos.vpCodigo1
         MSFlex.TextMatrix(k, 8) = frmLogVehiculoDatos.vpCodigo2
      Else
         MSFlex.TextMatrix(k, 7) = GetUbigeoCorto(frmLogVehiculoDatos.vpCodigo1)
         MSFlex.TextMatrix(k, 8) = GetUbigeoCorto(frmLogVehiculoDatos.vpCodigo2)
      End If
      MSFlex.TextMatrix(k, 10) = nAsignaNro
   End If
   
   cmdAgregar.SetFocus
End If
Exit Sub

Sal_Agregar:

   If HayTrans Then
      oConn.RollbackTrans
   End If
End Sub

'--------------------------------------------------------------------------------

Sub RecuperaDatos(pnAsignacionNro As Integer)
Dim oConn As New DConecta, sSQL As String
Dim rs As New ADODB.Recordset, k As Integer
Dim DLog As New DLogVehiculos

If oConn.AbreConexion Then
  
'   sSQL = "select m.nTipoReg,m.cValor0,m.cValor1,m.cValor2,t.cTipoReg as cRegistro," & _
'          "       m.cDescripcion,cDesc1=coalesce(o.cUbigeoDescripcion,''), cDesc2=coalesce(d.cUbigeoDescripcion,''), m.nMonto " & _
'          "  from LogVehiculoMovDet m " & _
'          " inner join (select nConsValor AS nTipoReg,cConsDescripcion as cTipoReg from Constante where nConsCod=9024 and nconscod<>nconsvalor) t on m.nTipoReg=t.nTipoReg " & _
'          " left outer join UbicacionGeografica o on m.cValor1=o.cUbigeoCod " & _
'          " left outer join UbicacionGeografica d on m.cValor2=d.cUbigeoCod " & _
'          " where m.nAsignacionNro = '" & pnAsignacionNro & "' and m.nEstado = 1"
'
    Set rs = DLog.GetVehiculoMovDet(pnAsignacionNro)
   If Not rs.EOF Then
      k = 0
      Do While Not rs.EOF
         k = k + 1
         InsRow MSFlex, k
         MSFlex.TextMatrix(k, 0) = rs!dFecha
         MSFlex.TextMatrix(k, 1) = rs!nTipoReg
         MSFlex.TextMatrix(k, 2) = rs!cValor0
         MSFlex.TextMatrix(k, 3) = rs!cValor1
         MSFlex.TextMatrix(k, 4) = rs!cValor2
         MSFlex.TextMatrix(k, 6) = rs!cDescripcion
         MSFlex.TextMatrix(k, 9) = FNumero(rs!nMonto)
         MSFlex.TextMatrix(k, 5) = rs!cRegistro
         If rs!nTipoReg = 4 Then
            MSFlex.TextMatrix(k, 7) = rs!cValor1
            MSFlex.TextMatrix(k, 8) = rs!cValor2
         Else
            MSFlex.TextMatrix(k, 7) = rs!cDesc1
            MSFlex.TextMatrix(k, 8) = rs!cDesc2
         End If
         MSFlex.TextMatrix(k, 10) = pnAsignacionNro
         rs.MoveNext
      Loop
   End If
End If
End Sub


Private Sub cmdQuitar_Click()
Dim oConn As New DConecta
Dim i As Integer, n As Integer
Dim cMovNro As String
Dim nTipoReg As Integer
Dim sSQL As String

If Len(MSFlex.TextMatrix(MSFlex.row, 10)) = 0 Then Exit Sub

cMovNro = MSFlex.TextMatrix(MSFlex.row, 10)
nTipoReg = MSFlex.TextMatrix(MSFlex.row, 1)

n = MSFlex.Rows - 1
If n > 1 Then
   MSFlex.RemoveItem MSFlex.row
Else
   For i = 0 To MSFlex.Cols - 1
       MSFlex.TextMatrix(n, i) = ""
   Next
   MSFlex.RowHeight(n) = 8
End If

sSQL = "UPDATE LogVehiculoAsignacionDet SET nEstado=0 Where nAsignacionNro = '" & nAsignaNro & "' and nTipoReg = " & nTipoReg & " "
If oConn.AbreConexion Then
   oConn.Ejecutar sSQL
   oConn.CierraConexion
End If
End Sub



