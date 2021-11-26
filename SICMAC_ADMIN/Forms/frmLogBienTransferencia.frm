VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLogBienTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Activo Fijo y Bienes no Depreciables"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "frmLogBienTransferencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEstadistico 
      Caption         =   "Solo estadístico / No genera asiento contable"
      Height          =   435
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "&Transferir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   5925
      TabIndex        =   9
      Top             =   3075
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   7000
      TabIndex        =   10
      Top             =   3075
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   40
      TabIndex        =   19
      Top             =   1680
      Width           =   8010
      Begin VB.CheckBox chkDestAreaAgencia 
         Caption         =   "Misma Área/Agencia"
         Height          =   195
         Left            =   6070
         TabIndex        =   8
         Top             =   280
         Width           =   1815
      End
      Begin VB.TextBox txtDestGlosa 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   6615
      End
      Begin VB.TextBox txtDestAreaAgeNombre 
         Height          =   285
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtDestUsuarioNombre 
         Height          =   285
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   2775
      End
      Begin Sicmact.TxtBuscar txtDestAreaAgeCod 
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Sicmact.TxtBuscar txtDestUsuarioCod 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
      End
      Begin MSComCtl2.DTPicker txtDestFechaTransf 
         Height          =   315
         Left            =   6670
         TabIndex        =   6
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   137625601
         CurrentDate     =   41414
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   5880
         TabIndex        =   25
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Área/Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   630
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   40
      TabIndex        =   11
      Top             =   0
      Width           =   8010
      Begin VB.CheckBox chkTipoBus 
         Caption         =   "Busqueda Por Serie"
         Height          =   435
         Left            =   6720
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtOrgSerieNombre 
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtOrgUsuarioCod 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtOrgUsuarioNombre 
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtOrgBienNombre 
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtOrgAreaAgeNombre 
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
      Begin Sicmact.TxtBuscar txtOrgAreaAgeCod 
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   450
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Sicmact.TxtBuscar txtOrgBienCod 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   450
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtBuscaSerie 
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin Sicmact.TxtBuscar txtOrgSerieCod 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   450
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   6
         EnabledText     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Serie:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   970
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Bien:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   625
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Área/Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1095
      End
   End
   Begin Sicmact.Usuario oUser 
      Left            =   7800
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmLogBienTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogBienTransferencia
'** Descripción : Transferencia de Bienes creado segun ERS059-2013
'** Creación : EJVG, 20130618 03:30:00 AM
'***************************************************************************
Option Explicit
Dim fbUnico As Boolean
Dim fnMovNro As Long, fnId As Long
Dim fsSerieCod As String, fsSerieNombre As String
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub chkTipoBus_Click()
    If chkTipoBus.value = 1 Then
        'MsgBox "chequeado"
        
        txtOrgAreaAgeCod.Enabled = False
        txtOrgBienCod.Enabled = False
        txtOrgSerieCod.Enabled = False
        
        txtBuscaSerie.Visible = True
        txtOrgSerieCod.Visible = False
        
        txtBuscaSerie.SetFocus
        
    Else
        'MsgBox "deschequeado"
        
        txtOrgAreaAgeCod.Enabled = True
        txtOrgBienCod.Enabled = True
        txtOrgSerieCod.Enabled = True
        
        txtBuscaSerie.Visible = False
        txtOrgSerieCod.Visible = True
        
    End If
    
    
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargaControles
    Limpiar
End Sub
Private Sub Limpiar()
    txtOrgAreaAgeCod.Text = ""
    txtOrgAreaAgeNombre.Text = ""
    txtOrgBienCod.Text = ""
    txtOrgBienNombre.Text = ""
    txtOrgSerieCod.Text = ""
    txtOrgSerieNombre.Text = ""
    txtOrgUsuarioCod.Text = ""
    txtOrgUsuarioNombre.Text = ""
    txtDestAreaAgeCod.Text = ""
    txtDestAreaAgeNombre.Text = ""
    txtDestUsuarioCod.Text = ""
    txtDestUsuarioNombre.Text = ""
    txtDestGlosa.Text = ""
    chkDestAreaAgencia.value = 0
    'chkEstadistico.value = 0
    
    'Me.chkTipoBus.value = 0 '*** PEAC 20140611
    Me.txtBuscaSerie.Text = "" '*** PEAC 20140611
    
    txtDestFechaTransf.value = Format(gdFecSis, gsFormatoFechaView)
    chkEstadistico.value = 0 'PASI20150318
End Sub
Private Sub CargaControles()
    Dim oArea As New DActualizaDatosArea
    txtOrgAreaAgeCod.rs = oArea.GetAgenciasAreas()
    txtDestAreaAgeCod.rs = oArea.GetAgenciasAreas()
    Set oArea = Nothing
End Sub

Private Sub txtBuscaSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtBuscaSerie.Text = "" Then
            Exit Sub
        End If
    
        Dim oBien As New DBien
        Dim lsDesc As String
        Dim Mat As Variant
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
    
                Set rs = oBien.RecuperaAreaAgenciaBienParaTransf(Trim(Me.txtBuscaSerie.Text))
    
                If rs.RecordCount = 0 Then
                    MsgBox "Número de serie ingresado NO existe.", vbOKOnly + vbExclamation, "Atención"
                    Me.txtBuscaSerie.SetFocus
                    Exit Sub
                End If
    
                txtOrgSerieNombre.Text = ""
                fsSerieCod = ""
                fsSerieNombre = ""
    
                txtOrgSerieCod.psDescripcion = rs!cNombre
    
                txtOrgSerieCod.psDescripcion = txtOrgSerieCod.psDescripcion & Space(500) & Trim(Str(rs!nUnico)) & "," & Trim(Str(rs!nMovNro)) & "," & Trim(Str(rs!nId)) & "," & Trim(rs!cPersCod) & "," & Trim(rs!cPersNombre)
    
                fsSerieCod = Me.txtBuscaSerie.Text
                fsSerieNombre = txtOrgSerieCod.psDescripcion
                txtOrgSerieNombre.Text = Trim(Left(txtOrgSerieCod.psDescripcion, 500))
    
                lsDesc = Trim(Right(txtOrgSerieCod.psDescripcion, 500))
    
                Mat = Split(lsDesc, ",")
    
                fbUnico = IIf(Mat(0) = 1, True, False)
                fnMovNro = Mat(1)
                fnId = Mat(2)
                txtOrgUsuarioCod.Text = Mat(3)
                txtOrgUsuarioNombre.Text = Mat(4)
    
                txtOrgAreaAgeCod.Text = IIf(rs!ccodareaagencia = "", rs!cAreaCod, rs!cAreaCod + rs!cAgeCod)
                txtOrgBienCod.Text = rs!cBien
                Me.txtOrgAreaAgeNombre.Text = IIf(rs!ccodareaagencia = "", rs!cNomArea, rs!cNomAge)
                Me.txtOrgBienNombre.Text = rs!cNomBien
                
                txtDestAreaAgeCod.SetFocus
                
    End If
End Sub

Private Sub txtDestGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdTransferir.SetFocus
    End If
End Sub
Private Sub txtDestUsuarioCod_EmiteDatos()
    txtDestUsuarioNombre.Text = ""
    If txtDestUsuarioCod.Text <> "" Then
        txtDestUsuarioNombre.Text = txtDestUsuarioCod.psDescripcion
    End If
End Sub
Private Sub txtOrgAreaAgeCod_EmiteDatos()
    Dim oBien As New DBien
    Screen.MousePointer = 11
    txtOrgAreaAgeNombre.Text = ""
    If txtOrgAreaAgeCod.Text <> "" Then
        txtOrgAreaAgeNombre.Text = txtOrgAreaAgeCod.psDescripcion
    End If
    txtOrgBienCod.rs = oBien.RecuperaCategoriasBienPaObjeto(False, Left(txtOrgAreaAgeCod.Text, 3) & IIf(Mid(txtOrgAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtOrgAreaAgeCod.Text, 4, 2)))
    txtOrgBienCod.Text = ""
    Call txtOrgBienCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtOrgAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOrgBienCod.SetFocus
    End If
End Sub
Private Sub txtOrgAreaAgeCod_LostFocus()
    If txtOrgAreaAgeCod.Text = "" Then
        txtOrgAreaAgeNombre.Text = ""
    End If
End Sub
Private Sub txtOrgBienCod_GotFocus()
'    If txtOrgAreaAgeCod.Text = "" Then
'        MsgBox "Ud. debe de seleccionar primero el Área/Agencia del Bien a Transferir", vbInformation, "Aviso"
'        txtOrgAreaAgeCod.SetFocus
'        Exit Sub
'    End If
End Sub
Private Sub txtOrgBienCod_EmiteDatos()
    txtOrgBienNombre.Text = ""
    If txtOrgBienCod.Text <> "" Then
        txtOrgBienNombre.Text = txtOrgBienCod.psDescripcion
    End If
    txtOrgSerieCod.Text = ""
    Call txtOrgSerieCod_EmiteDatos
End Sub
Private Sub txtOrgBienCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOrgSerieCod.SetFocus
    End If
End Sub
Private Sub txtOrgBienCod_LostFocus()
    If txtOrgBienCod.Text = "" Then
        txtOrgBienNombre.Text = ""
    End If
End Sub
Private Sub txtOrgSerieCod_GotFocus()
'*** PEAC 20140527
'    If txtOrgAreaAgeCod.Text = "" Then
'        MsgBox "Ud. debe de seleccionar primero el Área/Agencia del Bien a Transferir", vbInformation, "Aviso"
'        txtOrgAreaAgeCod.SetFocus
'        Exit Sub
'    End If
'    If txtOrgBienCod.Text = "" Then
'        MsgBox "Ud. debe de seleccionar primero la Categoría del Bien a Transferir", vbInformation, "Aviso"
'        txtOrgBienCod.SetFocus
'        Exit Sub
'    End If
End Sub
Private Sub txtOrgSerieCod_Click(psCodigo As String, psDescripcion As String)
    Call frmLogBienTransferenciaSel.Inicio(fsSerieCod, fsSerieNombre, Left(txtOrgAreaAgeCod.Text, 3) & IIf(Mid(txtOrgAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtOrgAreaAgeCod.Text, 4, 2)), txtOrgBienCod.Text)
    psCodigo = fsSerieCod
    psDescripcion = fsSerieNombre
End Sub

Private Sub txtOrgSerieCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsDesc As String
    Dim Mat As Variant
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

        txtOrgSerieNombre.Text = ""
        fsSerieCod = ""
        fsSerieNombre = ""
        If txtOrgSerieCod.Text <> "" Then
                
            fsSerieCod = txtOrgSerieCod.Text
            fsSerieNombre = txtOrgSerieCod.psDescripcion
            txtOrgSerieNombre.Text = Trim(Left(txtOrgSerieCod.psDescripcion, 500))
            lsDesc = Trim(Right(txtOrgSerieCod.psDescripcion, 500))
            Mat = Split(lsDesc, ",")
            fbUnico = IIf(Mat(0) = 1, True, False)
            fnMovNro = Mat(1)
            fnId = Mat(2)
            txtOrgUsuarioCod.Text = Mat(3)
            txtOrgUsuarioNombre.Text = Mat(4)
             
        End If

End Sub
Private Sub txtDestAreaAgeCod_EmiteDatos()
    Dim lcCodAgeDes As String
    txtDestAreaAgeNombre.Text = ""

    If txtDestAreaAgeCod.Text <> "" Then
        txtDestAreaAgeNombre.Text = IIf(Len(txtDestAreaAgeCod.psDescripcion) = 0, txtOrgAreaAgeNombre.Text, txtDestAreaAgeCod.psDescripcion) '*** PEAC 20140611
    End If
End Sub
Private Sub txtDestAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDestUsuarioCod.SetFocus
    End If
End Sub
Private Sub txtDestFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDestGlosa.SetFocus
    End If
End Sub
Private Sub chkDestAreaAgencia_Click()
    txtDestAreaAgeCod.Text = ""
    txtDestAreaAgeNombre.Text = ""
    If chkDestAreaAgencia.value = 1 Then
        txtDestAreaAgeCod.Text = txtOrgAreaAgeCod.Text
        Call txtDestAreaAgeCod_EmiteDatos
    End If
End Sub
Private Sub txtDestUsuarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDestGlosa.SetFocus
    End If
End Sub
Private Sub cmdCancelar_Click()
    Limpiar
End Sub
Private Sub cmdTransferir_Click()
    Dim oBien As DBien
    Dim oMov As DMov
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Dim ofun As NContFunciones
    Dim rsDet As ADODB.Recordset

    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim bTransBien As Boolean, bTransMov As Boolean
    Dim lnMovItem As Long
    Dim lsPlantilla1 As String, lsPlantilla2 As String
    Dim lsCtaContOrig As String, lsCtaContDest As String
    
    Dim lsOrigAreaCod As String, lsOrigAgeCod As String, lsOrigPersCod As String, lsOrigSerie As String
    Dim lsDestAreaCod As String, lsDestAgeCod As String, lsDestPersCod As String, lsDestGlosa As String
    Dim ldDestFechaTransf As Date
    Dim lsMovNroTotal As String
    Dim Movs As Variant
    Dim I As Integer
    Dim iSleep As Integer
    Dim lsImpre As String
    Dim lbMismaAgencia As Boolean
    
    If Not validaTransferir Then Exit Sub
    Set ofun = New NContFunciones
    On Error GoTo ErrTransferir
    lsOrigAreaCod = Left(txtOrgAreaAgeCod.Text, 3)
    lsOrigAgeCod = IIf(Mid(txtOrgAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtOrgAreaAgeCod.Text, 4, 2))
    lsOrigPersCod = Trim(txtOrgUsuarioCod.Text)
    lsDestAreaCod = Left(txtDestAreaAgeCod.Text, 3)
    lsDestAgeCod = IIf(Mid(txtDestAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtDestAreaAgeCod.Text, 4, 2))
    lsDestPersCod = txtDestUsuarioCod.Text
    ldDestFechaTransf = CDate(txtDestFechaTransf.value)
    lsDestGlosa = Trim(txtDestGlosa.Text)
    lsPlantilla1 = "18MRUAG"
    lsPlantilla2 = "18M90RUAG"
    
    If lsOrigAgeCod = lsDestAgeCod Then
        lbMismaAgencia = True
    End If

    If Not ofun.PermiteModificarAsiento(Format(ldDestFechaTransf, gsFormatoMovFecha), False) Then
        MsgBox "No se podrá realizar el Proceso ya que la fecha de transferencia pertenece a un mes ya cerrado", vbInformation, "Aviso"
        txtDestFechaTransf.SetFocus
        Set ofun = Nothing
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de realizar la Transferencia del Bien?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Set oBien = New DBien
    Set rsDet = oBien.RecuperaBienxTransferencia(fbUnico, fnMovNro, fnId)
    If rsDet.RecordCount = 0 Then
        MsgBox "No existe datos para realizar la Transferencia", vbCritical, "Aviso"
        Exit Sub
    End If

    Set oMov = New DMov
    
    oMov.BeginTrans
    oBien.dBeginTrans
    bTransMov = True
    bTransBien = True
    
    'Solo se va a generar asientos contables si es Activo Fijo
    Do While Not rsDet.EOF
        If ldDestFechaTransf < rsDet!dActivacion Then
            MsgBox "No se podrá realizar el Proceso ya que la fecha de transferencia es menor que la de Activación", vbInformation, "Aviso"
            oMov.RollbackTrans
            oBien.dRollbackTrans
            Set oMov = Nothing
            Set oBien = Nothing
            txtDestFechaTransf.SetFocus
            Exit Sub
        End If

        For iSleep = 0 To Rnd(2000) * 1000
        Next

        lnMovItem = 0
        lsMovNro = oMov.GeneraMovNro(ldDestFechaTransf, Right(gsCodAge, 2), gsCodUser)
        lsMovNroTotal = lsMovNroTotal & lsMovNro & ","
        oMov.InsertaMov lsMovNro, gnTransAF, Left(lsDestGlosa, 250), gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovBSAF Year(rsDet!dActivacion), rsDet!nMovNro, 1, rsDet!cObjetoCod, fsSerieCod, lnMovNro
        
        lnMovItem = lnMovItem + 1
        oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjCMACAgenciaArea
        oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, 1, lsOrigAgeCod, lsOrigAreaCod
        oMov.InsertaMovObjPers lnMovNro, lnMovItem, 1, lsOrigPersCod
        lsCtaContOrig = ReemplazaPlantilla(lsPlantilla1, lsOrigAgeCod, rsDet!nMoneda, rsDet!nBANCod)
        
        If Me.chkEstadistico.value = 0 Then 'PASI20150318
            If Not lbMismaAgencia And rsDet!cCategoBien = 1 Then
                oMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContOrig, (rsDet!nMontoBien) * -1
            End If
        End If
        lsCtaContOrig = ReemplazaPlantilla(lsPlantilla2, lsOrigAgeCod, rsDet!nMoneda, rsDet!nBANCod)
        lnMovItem = lnMovItem + 1
        
        If Me.chkEstadistico.value = 0 Then 'PASI20150318
            If Not lbMismaAgencia And rsDet!cCategoBien = 1 Then
                oMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContOrig, (rsDet!nDepreAcumulada)
            End If
        End If
        
        oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjCMACAgenciaArea
        oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, 1, lsDestAgeCod, lsDestAreaCod
        oMov.InsertaMovObjPers lnMovNro, lnMovItem, 1, lsDestPersCod
        lsCtaContDest = ReemplazaPlantilla(lsPlantilla1, lsDestAgeCod, rsDet!nMoneda, rsDet!nBANCod)
        lnMovItem = lnMovItem + 1
        
        If Me.chkEstadistico.value = 0 Then 'PASI20150318
            If Not lbMismaAgencia And rsDet!cCategoBien = 1 Then
                oMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContDest, rsDet!nMontoBien
            End If
        End If
        
        lsCtaContDest = ReemplazaPlantilla(lsPlantilla2, lsDestAgeCod, rsDet!nMoneda, rsDet!nBANCod)
        lnMovItem = lnMovItem + 1
        
        If Me.chkEstadistico.value = 0 Then 'PASI20150318
            If Not lbMismaAgencia And rsDet!cCategoBien = 1 Then
                oMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaContDest, (rsDet!nDepreAcumulada) * -1
            End If
        End If
        
        If Me.chkEstadistico.value = 0 Then 'PASI20150318
            If rsDet!nMoneda = 2 Then
                oMov.GeneraMovME lnMovNro, lsMovNro
            End If
        End If
        Call oBien.ActualizarAF(rsDet!nMovNro, , , , , , lsDestAreaCod, lsDestAgeCod, lsDestPersCod)
        rsDet.MoveNext
    Loop
    If fbUnico = False Then
        Call oBien.ActualizarActivoCompuesto(fnId, lsDestAreaCod, lsDestAgeCod, lsDestPersCod)
    End If
    
    oMov.CommitTrans
    oBien.dCommitTrans
    bTransMov = False
    bTransBien = False
    
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Movs = Split(lsMovNroTotal, ",")
    
    Screen.MousePointer = 0
    MsgBox "Se ha realizado el proceso de Transferencia satisfactoriamente", vbInformation, "Aviso"
    If Me.chkEstadistico.value = 0 Then 'PASI20150318
        For I = 0 To UBound(Movs) - 1
            lsImpre = oAsiento.ImprimeAsientoContable(Movs(I), 60, 80, Caption)
            If lsImpre <> "" Then
                oPrevio.Show lsImpre, Caption, True
            End If
        Next
    End If
    
        'ARLO 20160126 ***
        'gsopecod = LogPistaReporteEstadistico
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se ha realizado el proceso de Transferencia satisfactoriamente de la Agencia Origen : " & txtOrgAreaAgeNombre & " el Bien " & txtOrgBienNombre & " a la Agencia destino " & txtDestAreaAgeNombre
        Set objPista = Nothing
        '**************
    Limpiar
    Set oBien = Nothing
    Set oMov = Nothing
    Set ofun = Nothing
    Set oPrevio = Nothing
    Set oAsiento = Nothing

        
    Exit Sub
ErrTransferir:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTransBien Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    If bTransMov Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
End Sub
Private Function ReemplazaPlantilla(ByVal psPlantilla As String, ByVal psAgeCod As String, ByVal pnMoneda As Integer, ByVal pnBANCod As Integer) As String
    Dim obj As New DBien
    Dim lsRubro As String
    lsRubro = obj.RecuperaSubCtaxBANCod(pnBANCod)
    ReemplazaPlantilla = psPlantilla
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "M", pnMoneda)
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "RU", lsRubro)
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "AG", psAgeCod)
    Set obj = Nothing
End Function
Private Function validaTransferir() As Boolean
    Dim lsValFecha As String
    validaTransferir = True
    If Len(Trim(txtOrgAreaAgeCod.Text)) = 0 Then
        validaTransferir = False
        MsgBox "Ud. debe de seleccionar el Área/Agencia del Bien a Transferir", vbInformation, "Aviso"
        txtOrgAreaAgeCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtOrgBienCod.Text)) = 0 Then
        validaTransferir = False
        MsgBox "Ud. debe de seleccionar la Categoría del Bien a Transferir", vbInformation, "Aviso"
        txtOrgBienCod.SetFocus
        Exit Function
    End If
    
    If Me.txtOrgSerieCod.Visible = True Then
        If Len(Trim(txtOrgSerieCod.Text)) = 0 Then
            validaTransferir = False
            MsgBox "Ud. debe de seleccionar el Bien a Transferir", vbInformation, "Aviso"
            txtOrgSerieCod.SetFocus
        End If
        Exit Function
    Else
        If Len(Trim(Me.txtBuscaSerie.Text)) = 0 Then
            validaTransferir = False
            MsgBox "Ud. debe de seleccionar el Bien a Transferir", vbInformation, "Aviso"
            Me.txtBuscaSerie.SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtDestAreaAgeCod.Text)) = 0 Then
        validaTransferir = False
        MsgBox "Ud. debe de seleccionar el Área/Agencia Destino", vbInformation, "Aviso"
        txtDestAreaAgeCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDestUsuarioCod.Text)) = 0 Then
        validaTransferir = False
        MsgBox "Ud. debe de seleccionar el Usuario Destino", vbInformation, "Aviso"
        txtDestUsuarioCod.SetFocus
        Exit Function
    Else
        oUser.DatosPers (txtDestUsuarioCod.Text)
        If oUser.AreaCod <> Left(txtDestAreaAgeCod.Text, 3) Then
            validaTransferir = False
            MsgBox "La Persona ingresada debe pertenecer al Área seleccionada", vbInformation, "Aviso"
            txtDestUsuarioCod.SetFocus
            Exit Function
        End If
    End If
    lsValFecha = ValidaFecha(txtDestFechaTransf.value)
    If Len(lsValFecha) > 0 Then
        validaTransferir = False
        MsgBox lsValFecha, vbInformation, "Aviso"
        txtDestFechaTransf.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDestGlosa.Text)) = 0 Then
        validaTransferir = False
        MsgBox "Ud. debe de ingresar la glosa respectiva", vbInformation, "Aviso"
        txtDestGlosa.SetFocus
        Exit Function
    End If
End Function

Private Sub txtOrgSerieCod_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtDestAreaAgeCod.SetFocus
'    End If
End Sub

Private Sub txtOrgSerieCod_LostFocus()
'    If txtOrgSerieCod.Text = "" Then
'        txtOrgSerieCod.Text = ""
'    End If
End Sub


