VERSION 5.00
Begin VB.Form frmlogreqrechazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio De Estado"
   ClientHeight    =   1650
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "frmlogreqrechazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbestado 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   225
      Width           =   2415
   End
   Begin VB.ComboBox cmbmotrechazo 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Requerimiento"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label lblmotivorechazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo del Rechazo"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmlogreqrechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDGnral As DLogGeneral
Dim clsDMov As DLogMov
Public bseleccion As Boolean
Option Explicit

Private Sub CancelButton_Click()
bseleccion = False
Unload Me
End Sub


Private Sub cmbestado_Click()

Select Case Right(Trim(cmbestado.Text), 1)
Case "0" 'seleccione
        cmbmotrechazo.Visible = False
        lblmotivorechazo.Visible = False
Case "1" 'Solicitados
        cmbmotrechazo.Visible = False
        lblmotivorechazo.Visible = False
Case "2" 'Aprobados
        cmbmotrechazo.Visible = False
        lblmotivorechazo.Visible = False
Case "3" 'Rechazados
        cmbmotrechazo.Visible = True
        lblmotivorechazo.Visible = True
End Select
End Sub


Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim nestado As Integer
Call CentraForm(Me)
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set rs = clsDGnral.CargaMotRechazos
Call CargaCombo(rs, cmbmotrechazo)
rs.Close
bseleccion = False

'Obtener el estado del requerimiento
nestado = clsDGnral.GetEstadoRequerimiento(Trim(frmLogReqInicio.TxtBuscar))
Select Case nestado
Case 1 'solicitado
        cmbestado.AddItem "Aprobado  " & Space(100) & "2"
        cmbestado.AddItem "Rechazado" & Space(100) & "3"
Case 2 'Aprobado
        cmbestado.AddItem "Solicitado" & Space(100) & "1"
        cmbestado.AddItem "Rechazado" & Space(100) & "3"
Case 3 'Rechazado
        cmbestado.AddItem "Aprobado" & Space(100) & "2"
        cmbestado.AddItem "Solicitado" & Space(100) & "1"
End Select


End Sub


Private Sub OKButton_Click()
Dim sReqTraNro As String
Dim sactualiza As String
Dim nReqTraNro  As String
Dim sReqNro  As String
Dim nReqNro   As Long
Set clsDGnral = New DLogGeneral
Set clsDMov = New DLogMov
    If cmbestado.Text = "" Then
        MsgBox "Antes debe Seleccionar un Estado  para el Requerimiento ", vbInformation, "Seleccione Un Estado"
        Exit Sub
    End If
Select Case Right(Trim(cmbestado.Text), 1)
Case 1
     If MsgBox("Desea  Cambiar el Estado a Requeimiento Solcitado ?   ", vbQuestion + vbYesNo, " Cambio de estado ---- > Solicitado") = vbYes Then
            sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqModifica)), "", gLogReqEstadoInicio
            nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
            'nReqNro = clsDMov.GetnMovNro(sReqNro)
            'clsDMov.InsertaMovRef nReqTraNro, nReqNro
           clsDGnral.ActualizaEstReq Trim(frmLogReqInicio.TxtBuscar.Text), "", ReqEstadoSolicitado, "A"
            Unload Me
     End If
Case 2
    If MsgBox("Desea  Aprobar el Requerimiento   ", vbQuestion + vbYesNo, "  Cambio de estado ---- > Aprobacion") = vbYes Then
            sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqModifica)), "", gLogReqEstadoInicio
            nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
            'nReqNro = clsDMov.GetnMovNro(sReqNro)
            'clsDMov.InsertaMovRef nReqTraNro, nReqNro
            clsDGnral.ActualizaEstReq Trim(frmLogReqInicio.TxtBuscar.Text), "", ReqEstadoaprobado, "A"
            Unload Me
    End If
Case 3
        If cmbmotrechazo.Text = "" Then
            MsgBox "Antes debe Seleccionar un tipo de Rechazo ", vbInformation, "Seleccione Un tipo de Rechazo"
            Exit Sub
            cmbmotrechazo.SetFocus
        End If
  If MsgBox("Desea  Rechazar el Requerimiento ?  ", vbQuestion + vbYesNo, " Cambio de estado ---- > Rechazado") = vbYes Then
            sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqModifica)), "", gLogReqEstadoInicio
            nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
            'nReqNro = clsDMov.GetnMovNro(sReqNro)
            'clsDMov.InsertaMovRef nReqTraNro, nReqNro
            clsDGnral.ActualizaEstReq Trim(frmLogReqInicio.TxtBuscar.Text), Right(Trim(cmbmotrechazo.Text), 2), ReqEstadoObservado, "R"
            Unload Me
    End If
End Select
bseleccion = True
'clsDMov.ActualizaRequeri Trim(txtBuscar.Text), cbomePeriodo.Text, IIf(psTpoReq = "1", gLogReqTipoNormal, gLogReqTipoExtemporaneo), rtfDescri(0).Text, rtfDescri(1).Text, Trim(Txtarea)
bseleccion = True
Unload Me
End Sub
