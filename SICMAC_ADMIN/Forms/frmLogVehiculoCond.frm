VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoCond 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Conductores"
   ClientHeight    =   3420
   ClientLeft      =   1635
   ClientTop       =   3045
   ClientWidth     =   8535
   Icon            =   "frmLogVehiculoCond.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVis 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   370
         Left            =   7080
         TabIndex        =   3
         Top             =   3000
         Width           =   1200
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   370
         Left            =   0
         TabIndex        =   2
         Top             =   3000
         Width           =   1200
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   1260
         TabIndex        =   1
         Top             =   3000
         Width           =   1140
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   2715
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4789
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
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
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame fraReg 
      Height          =   3150
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton CmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2145
         TabIndex        =   13
         Top             =   510
         Width           =   375
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6840
         TabIndex        =   12
         Top             =   2580
         Width           =   1140
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5580
         TabIndex        =   11
         Top             =   2580
         Width           =   1200
      End
      Begin VB.TextBox txtPersona 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   5475
      End
      Begin VB.TextBox txtBrevete 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   1515
      End
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   6975
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   6975
      End
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   6975
      End
      Begin VB.TextBox txtPersCod 
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
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Brevete Nº"
         Height          =   195
         Left            =   5580
         TabIndex        =   18
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1620
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmLogVehiculoCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cCargoCod As String, cAreaCod As String, cAgenciaCod As String, cBrevete As String

Private Sub cmdAgregar_Click()
fraVis.Visible = False
fraReg.Visible = True
txtPersCod.Text = ""
txtPersona.Text = ""
cCargoCod = ""
cAreaCod = ""
cAgenciaCod = ""
txtCargo.Text = ""
txtArea.Text = ""
cboAgencia.ListIndex = -1
End Sub

Private Sub cmdQuitar_Click()
If MsgBox("¿ Está seguro de quitar al Conductor indicado ? " + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
  
End If
End Sub

Private Sub Form_Load()
CentraForm Me
CargaAgencias
CargaFlex
End Sub

Sub CargaAgencias()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim sSQL As String
If oConn.AbreConexion Then
   sSQL = "Select cAgeCod, cAgeDescripcion from Agencias where nEstado=1"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboAgencia.AddItem rs!cAgeDescripcion
         cboAgencia.ItemData(cboAgencia.ListCount - 1) = rs!cAgeCod
         rs.MoveNext
      Loop
      cboAgencia.ListIndex = -1
   End If
End If
End Sub

Sub CargaFlex()
Dim LV As DLogVehiculos, rs As ADODB.Recordset
Dim i As Integer, Estado  As String

LimpiaFlex
Set LV = New DLogVehiculos
Set rs = LV.ListaConductores
i = 0
While Not rs.EOF
    i = i + 1
    InsRow Flex, i
    Flex.TextMatrix(i, 1) = rs!cPersCod
    Flex.TextMatrix(i, 2) = Replace(rs!cPersNombre, "/", " ")
    Flex.TextMatrix(i, 3) = rs!cBrevete
    Select Case rs!nEstado
        Case "0"
             Flex.TextMatrix(i, 4) = "Anulado" & Space(100) & CStr(rs!nEstado)
        Case "1"
             Flex.TextMatrix(i, 4) = "Disponible" & Space(100) & CStr(rs!nEstado)
        Case "2"
             Flex.TextMatrix(i, 4) = "Asignado" & Space(100) & CStr(rs!nEstado)
        Case "3"
             Flex.TextMatrix(i, 4) = "Vacaciones" & Space(100) & CStr(rs!nEstado)
        Case "4"
             Flex.TextMatrix(i, 4) = "Retirado" & Space(100) & CStr(rs!nEstado)
    End Select
    Flex.TextMatrix(i, 5) = rs!cAgencia
    'Flex.TextMatrix(i, 6) = rs!nEstado
    rs.MoveNext
Wend
Set rs = Nothing
Set LV = Nothing
End Sub

Sub LimpiaFlex()
Flex.Rows = 2
Flex.Clear
Flex.RowHeight(0) = 320
Flex.ColWidth(0) = 0
Flex.ColWidth(1) = 1100:  Flex.ColAlignment(1) = 4: Flex.TextMatrix(0, 1) = "    ID Persona"
Flex.ColWidth(2) = 2800:                            Flex.TextMatrix(0, 2) = "Persona"
Flex.ColWidth(3) = 1000:  Flex.ColAlignment(3) = 4: Flex.TextMatrix(0, 3) = "     Brevete"
Flex.ColWidth(4) = 1000:  Flex.ColAlignment(4) = 1: Flex.TextMatrix(0, 4) = " Estado"
Flex.ColWidth(5) = 2100
Flex.ColWidth(6) = 0
End Sub

Private Sub cmdAceptar_Click()
Dim opt As Integer
Dim v As DLogVehiculos

If Trim(txtPersCod) = "" Then Exit Sub
If Trim(txtBrevete) = "" Then
   If MsgBox("Es una persona sin brevete..." + vbCrLf + "¿ Grabar en esa condición ?", vbQuestion + vbYesNo, "Confirme") = vbNo Then
      Exit Sub
   End If
End If

Set v = New DLogVehiculos
If v.YaEstaRegistrado(txtPersCod.Text) Then
   MsgBox "La persona ya está registrada como Conductor..." + Space(10), vbInformation, "AVISO"
   cmdCancelar_Click
   Set v = Nothing
   Exit Sub
End If

If MsgBox("¿ Esta seguro de grabar ?" + Space(10), vbQuestion + vbYesNo, "AVISO") = vbYes Then
   cBrevete = txtBrevete
   If Len(cCargoCod) = 0 Then
      cAgenciaCod = Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")
   End If
   Call v.InsertRegConductor(txtPersCod, cCargoCod, cAreaCod, cAgenciaCod, cBrevete, 1, gdFecSis, Right(gsCodAge, 2), gsCodUser)
   CargaFlex
   cmdCancelar_Click
   Set v = Nothing
End If

End Sub

Private Sub cmdBuscar_Click()
Dim X As UPersona
Dim LV As DLogVehiculos
Dim oConn As New DConecta, rs As New ADODB.Recordset

cBrevete = ""
cCargoCod = ""
cAreaCod = ""
cAgenciaCod = ""

Set X = frmBuscaPersona.Inicio
If X Is Nothing Then
    Exit Sub
End If
'cPersCod = X.sPersCod

'frmBuscaPersona.Inicio True
'If frmBuscaPersona. Then
   txtPersona.Text = X.sPersNombre  'frmBuscaPersona.vpPersNom
   txtPersCod = X.sPersCod  'frmBuscaPersona.vpPersCod
   'Set rs = GetAreaCargoAgencia(txtPersCod)
   If Not rs.EOF Then
      cCargoCod = rs!cRHCargoCod
      cAreaCod = rs!cRHAreaCod
      cAgenciaCod = rs!cRHAgeCod
      txtCargo = rs!cRHCargo
      txtArea = rs!cRHArea
      txtCargo.BackColor = "&H80000005"
      txtArea.BackColor = "&H80000005"
      cboAgencia.Text = rs!cRHAgencia
   Else
      txtCargo = ""
      txtArea = ""
      txtCargo.BackColor = "&H8000000F"
      txtArea.BackColor = "&H8000000F"
      cboAgencia.ListIndex = -1
   End If
'End If


'If Len(X.sPersNombre) > 0 Then
If Len(txtPersCod) > 0 Then
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet("select cPersCod from LogVehiculoConductor where cPersCod = '" & txtPersCod & "'  ")
      If Not rs.EOF Then
         MsgBox "La persona ya está registrada como conductor..." + Space(10), vbInformation, "Aviso"
         txtPersona.Text = ""
         txtPersCod = ""
         Exit Sub
      End If
      oConn.CierraConexion
   End If
    
   Set LV = New DLogVehiculos
   cBrevete = LV.GetBrevete(txtPersCod)
   txtBrevete.Text = cBrevete
   If cBrevete = "" Then
      If MsgBox("Esta Persona no posee Brevete" + Space(10) + vbCrLf + "¿ Agregar como conductor ?" + Space(10), vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
         txtPersCod = ""
         txtPersona = ""
         txtCargo = ""
         txtArea = ""
         cboAgencia.ListIndex = -1
         Set LV = Nothing
         Exit Sub
      End If
    End If
End If
Set LV = Nothing

End Sub

Private Sub cmdCancelar_Click()
txtPersCod = ""
Me.txtPersona = ""
cBrevete = ""
'Me.lblBrevete.Text = ""
'Me.lblDNI = ""
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


