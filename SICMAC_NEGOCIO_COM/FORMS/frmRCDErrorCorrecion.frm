VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDErrorCorreccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe RCD - Correccion de Errores"
   ClientHeight    =   6300
   ClientLeft      =   75
   ClientTop       =   1065
   ClientWidth     =   11385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9900
      TabIndex        =   30
      Top             =   5820
      Width           =   1320
   End
   Begin VB.CommandButton cmdActualizaRRPP 
      Caption         =   "RRPP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdActualizaCIIU 
      Caption         =   "CIIU"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   28
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txthasta 
      Height          =   330
      Left            =   3255
      TabIndex        =   1
      Top             =   75
      Width           =   1380
   End
   Begin VB.TextBox txtDesde 
      Height          =   330
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   1380
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   9675
      TabIndex        =   2
      Top             =   60
      Width           =   1530
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualizar en base"
      Height          =   390
      Left            =   7740
      TabIndex        =   15
      Top             =   5820
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos a Actualizar"
      Height          =   1965
      Left            =   150
      TabIndex        =   17
      Top             =   3735
      Width           =   11130
      Begin VB.TextBox txtDocPersona 
         Height          =   300
         Left            =   9870
         TabIndex        =   26
         Top             =   705
         Width           =   1185
      End
      Begin VB.TextBox txtcodsbs 
         Height          =   345
         Left            =   9765
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1125
         Width           =   1305
      End
      Begin VB.TextBox txtcodsbsbase 
         Height          =   330
         Left            =   9765
         TabIndex        =   13
         Top             =   1485
         Width           =   1305
      End
      Begin VB.TextBox txtdatoMaestro 
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   2655
         TabIndex        =   9
         Top             =   630
         Width           =   6165
      End
      Begin VB.TextBox txtdatoBase 
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   2655
         TabIndex        =   12
         Top             =   1425
         Width           =   6120
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Ca&mbiar"
         Height          =   360
         Left            =   9600
         TabIndex        =   14
         Top             =   225
         Width           =   1470
      End
      Begin VB.TextBox txtdatoCambio 
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1020
         Width           =   6165
      End
      Begin VB.TextBox txtDato 
         Height          =   345
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   6150
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "Cod C&IIU"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   7
         Top             =   1545
         Width           =   1095
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "Cod &SBS"
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   6
         Top             =   1230
         Width           =   1395
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "&Cod RRPP"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   5
         Top             =   945
         Width           =   1095
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "&Documentos"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   615
         Width           =   1245
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "&Nombres"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Doc.Persona"
         Height          =   195
         Left            =   8880
         TabIndex        =   27
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SBS Base "
         Height          =   195
         Index           =   3
         Left            =   8925
         TabIndex        =   25
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod SBS  "
         Height          =   195
         Index           =   5
         Left            =   8925
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Dato Maestro :"
         Height          =   225
         Index           =   4
         Left            =   1440
         TabIndex        =   23
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Dato Base Persona :"
         Height          =   225
         Index           =   2
         Left            =   1440
         TabIndex        =   22
         Top             =   1455
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Datos SBS"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dato Rep."
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   20
         Top             =   285
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstErrores 
      Height          =   3210
      Left            =   105
      TabIndex        =   16
      Top             =   480
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5662
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Corr"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "cCodPers"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "cCodSBS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "cCodSBSBase"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre CMACT Rep"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nombre SBS"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nombre Base"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Nro Doc CMACT Rep"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "NroDoc SBS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "RRPP CMACT Rep"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "RRPP SBS"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CIIU CMACT Rep"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CIUU SBS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "CodUnico CMACT Rep"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "CodUnico SBS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "NombreMaestro"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Nro DocMaestro"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "RRPP Maestro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "CIIU Maestro"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta :"
      Height          =   195
      Index           =   1
      Left            =   2670
      TabIndex        =   19
      Top             =   105
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde :"
      Height          =   195
      Index           =   0
      Left            =   255
      TabIndex        =   18
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmRCDErrorCorreccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' RCD - Corrige errores enviados por SBS
'LAYG   :  10/01/2003.
'Resumen:  Nos permite corregir los errores enviados por la SBS

Option Explicit

Dim fsServConsol As String
Dim lnRegistros As Long


Private Sub cmdActualiza_Click()
Dim I As Integer
Dim lsSQL As String
Dim loBase As DConecta

If MsgBox("Desea Actualizar los datos??", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
Set loBase = New DConecta
    loBase.AbreConexion
    For I = 1 To Me.lstErrores.ListItems.Count
        If Me.lstErrores.ListItems(I).Checked = True Then
            If optActualiza(0).value = True Then  ' Nombre
                lsSQL = "UPDATE Persona set cPersNombre = '" & Replace(lstErrores.ListItems(I).SubItems(6), "'", "''") & "',  " _
                    & " cCodSBS = '" & lstErrores.ListItems(I).SubItems(3) & "' " _
                    & " WHERE cPersCod ='" & lstErrores.ListItems(I).SubItems(1) & "'"
                loBase.Ejecutar (lsSQL)
            
                lsSQL = "UPDATE RCDMaestroPersona set cPersNombre = '" & Replace(lstErrores.ListItems(I).SubItems(15), "'", "''") & "'  " _
                    & " WHERE cPersCod ='" & lstErrores.ListItems(I).SubItems(1) & "' "
                loBase.Ejecutar (lsSQL)
            End If
'            If optActualiza(1).Value = True Then  ' Documento
'                lsSQL = "UPDATE Persona set cNudoci='" & lstErrores.ListItems(I).SubItems(8) & "' " _
'                    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "'"
'                loBase.Ejecutar (lsSQL)
'
'                lsSQL = "UPDATE RCDMaestroPersona set cNudoci = '" & Replace(lstErrores.ListItems(I).SubItems(8), "'", "''") & "'  " _
'                    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "' "
'                loBase.Ejecutar (lsSQL)
'            End If
'
'            If optActualiza(2).Value = True Then  ' RRPP reg publicos
'                'sql = "UPDATE Persona set cCodSBS = '" & lstErrores.ListItems(I).SubItems(3) & "," _
'                '    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "'"
'                'dbCmact.Execute sql
'                lsSQL = "UPDATE RCDMaestroPersona set cCodRegPub = '" & Replace(lstErrores.ListItems(I).SubItems(10), "'", "''") & "'  " _
'                    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "' "
'                loBase.Ejecutar (lsSQL)
'            End If
'
'            If optActualiza(4).Value = True Then  ' CIIU
'                'sql = "UPDATE Persona set cCodSBS = '" & lstErrores.ListItems(I).SubItems(3) & "," _
'                '    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "'"
'                'dbCmact.Execute sql
'                lsSQL = "UPDATE RCDMaestroPersona set cActEcon = '" & Replace(lstErrores.ListItems(I).SubItems(12), "'", "''") & "'  " _
'                    & " WHERE cCodpers ='" & lstErrores.ListItems(I).SubItems(1) & "' "
'                loBase.Ejecutar (lsSQL)
'            End If
        
        End If
    Next

Set loBase = Nothing

End Sub

Private Sub cmdActualizaCIIU_Click()
Dim I As Integer
Dim sql As String
Dim loBase As DConecta
If MsgBox("Desea Actualizar CIIU en el RCDMaestroPersona ??", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
Set loBase = New DConecta
loBase.AbreConexion
For I = 1 To Me.lstErrores.ListItems.Count
    If lstErrores.ListItems(I).SubItems(12) <> "" Then
        sql = "UPDATE RCDMaestroPersona set cActEcon = '" & Replace(lstErrores.ListItems(I).SubItems(12), "'", "''") & "'  " _
            & " WHERE cCodPers ='" & lstErrores.ListItems(I).SubItems(1) & "' "
        
        loBase.Ejecutar sql
    End If
    
Next I
Set loBase = Nothing
MsgBox "Se ha actualizado correctamente el CIIU ", vbInformation, "Aviso"
End Sub

Private Sub cmdActualizaRRPP_Click()
Dim I As Integer
Dim sql As String
Dim loBase As DConecta
            
If MsgBox("Desea Actualizar RRPP en el RCDMaestroPersona ??", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
Set loBase = New DConecta
loBase.AbreConexion
For I = 1 To Me.lstErrores.ListItems.Count
    If lstErrores.ListItems(I).SubItems(12) <> "" Then
        sql = "UPDATE RCDMaestroPersona set cCodRegPub = '" & Replace(lstErrores.ListItems(I).SubItems(10), "'", "''") & "'  " _
            & " WHERE cCodPers ='" & lstErrores.ListItems(I).SubItems(1) & "' "
        
        loBase.Ejecutar sql
    End If
    
Next I
Set loBase = Nothing
MsgBox "Se ha actualizado correctamente el RRPP ", vbInformation, "Aviso"
End Sub

Private Sub cmdCambiar_Click()

If lnRegistros = 0 Then
    Exit Sub
End If

If Me.optActualiza(0).value = True Then ' nombres
    lstErrores.SelectedItem.SubItems(15) = Me.txtdatoMaestro
    lstErrores.SelectedItem.SubItems(6) = Me.txtdatoBase
    lstErrores.SelectedItem.SubItems(3) = txtcodsbsbase
End If
If Me.optActualiza(1).value = True Then  ' documentos
    lstErrores.SelectedItem.SubItems(8) = txtdatoBase
End If
If Me.optActualiza(2).value = True Then  'Cod RRPP
    lstErrores.SelectedItem.SubItems(10) = txtdatoBase
End If
If Me.optActualiza(4).value = True Then  'CIIU
    lstErrores.SelectedItem.SubItems(12) = txtdatoBase
End If
If Me.optActualiza(3).value = True Then 'codigos unicos
    lstErrores.SelectedItem.SubItems(3) = txtdatoBase
End If
lstErrores.SelectedItem.Checked = True
lstErrores.SetFocus
End Sub

Private Sub cmdProcesar_Click()
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim loBase As DConecta
Dim Item As ListItem

lstErrores.ListItems.Clear

lsSQL = "SELECT E.cCorr, E.cPersCod, E.cCodSBS, cNombreEmp,cNomSBS, ISNULL(P.cPersNombre,'') AS cNomPers, ISNULL(M.cCodSBS,'') as cCodSBSBase, " _
    & "       cNumDocEmp , cNumDocSBS, cCodRRPP, cCodRRPPSBS, cCodUnicoEmp, cCodUnicSBS , cCiiuEmp,cCiiuSBS , " _
    & "       Isnull(M.cPersNom,'') as PersonaMaestro, Isnull(m.cNuDoci,'') as NumDocMaestro, Isnull(m.cCodRegPub,'') as CodRegPubMaestro, " _
    & "       Isnull(M.cCodSBS,'') as CodSBSMaestro, Isnull(m.cActEcon,'') as cActEconMaestro " _
    & " FROM " & fsServConsol & "RCDError E  LEFT JOIN Persona P ON P.cPersCod =E.cPersCod " _
    & " LEFT JOIN " & fsServConsol & "RCDMaestroPersona M ON M.cPersCod=E.cPersCod  " _
    & " Where (Len(cNombreEmp) > 0 Or Len(cNumDocEmp) > 0 Or Len(cCodRRPP) > 0 Or cCodUnicoEmp > 0) " _
    & "         AND cCorr BETWEEN '" & txtDesde & "' AND '" & txthasta & "'" _
    & " ORDER BY cCorr "

Set loBase = New DConecta
    loBase.AbreConexion
    Set rs = loBase.CargaRecordSet(lsSQL)
Set loBase = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        lnRegistros = 1
        Do While Not rs.EOF
          If Trim(rs!cNomSBS) <> Trim(rs!personamaestro) Then
            Set Item = lstErrores.ListItems.Add(, , rs!cCorr)
            Item.SubItems(1) = Trim(rs!cCodPers)
            Item.SubItems(2) = Trim(rs!cCodSBS)
            Item.SubItems(3) = Trim(rs!cCodSBSBase)
            Item.SubItems(4) = Trim(rs!cNombreEmp)
            Item.SubItems(5) = Trim(rs!cNomSBS)
            Item.SubItems(6) = Trim(rs!cNomPers)
            Item.SubItems(7) = Trim(rs!cNumDocEmp)
            Item.SubItems(8) = Trim(rs!cNumDocSBS)
            Item.SubItems(9) = Trim(rs!cCodRRPP)
            Item.SubItems(10) = Trim(rs!cCodRRPPSBS)
            Item.SubItems(11) = Trim(rs!cCiiUEMP)
            Item.SubItems(12) = Trim(rs!cCiiuSBS)
            Item.SubItems(13) = Trim(rs!cCodUnicoEmp)
            Item.SubItems(14) = Trim(rs!cCodUnicSBS)
            Item.SubItems(15) = Trim(rs!personamaestro)
            Item.SubItems(16) = Trim(rs!NumDocMaestro)
            Item.SubItems(17) = Trim(rs!CodRegPubMaestro)
            'Item.SubItems(18) = Trim(rs!CodSBSMaestro)
            Item.SubItems(18) = Trim(rs!cActEconMaestro)
          End If
            
          DoEvents
          rs.MoveNext
        Loop
    Else
        lnRegistros = 0
    End If
    Set rs = Nothing
    
Me.cmdActualizaCIIU.Enabled = True
Me.cmdActualizaRRPP.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim loConstSistema As NConstSistemas
    
    Set loConstSistema = New NConstSistemas
        fsServConsol = loConstSistema.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSistema = Nothing
    
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub lstErrores_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtDato.Text = ""
txtdatoCambio = ""
txtdatoBase = ""
txtdatoMaestro = ""
txtcodsbs = ""
txtcodsbsbase = ""
txtDocPersona = lstErrores.SelectedItem.SubItems(16)

If Me.optActualiza(0).value = True Then ' nombres
    txtDato = lstErrores.SelectedItem.SubItems(4)
    txtdatoCambio = lstErrores.SelectedItem.SubItems(5)
    txtdatoBase = lstErrores.SelectedItem.SubItems(6)
    Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(15)
    txtcodsbs = lstErrores.SelectedItem.SubItems(2)
    If lstErrores.SelectedItem.SubItems(3) = "" Then
        txtcodsbsbase = lstErrores.SelectedItem.SubItems(2)
    Else
        txtcodsbsbase = lstErrores.SelectedItem.SubItems(3)
    End If
End If
If Me.optActualiza(1).value = True Then  ' documentos
    txtDato = lstErrores.SelectedItem.SubItems(7)
    txtdatoCambio = lstErrores.SelectedItem.SubItems(8)
    txtdatoBase = lstErrores.SelectedItem.SubItems(8)
    Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(16)
End If
If Me.optActualiza(2).value = True Then  'Cod RRPP
    txtDato = lstErrores.SelectedItem.SubItems(9)
    txtdatoCambio = lstErrores.SelectedItem.SubItems(10)
    txtdatoBase = lstErrores.SelectedItem.SubItems(10)
End If
If Me.optActualiza(3).value = True Then 'codigos unicos
    txtDato = lstErrores.SelectedItem.SubItems(2)
    txtdatoCambio = lstErrores.SelectedItem.SubItems(3)
    txtdatoBase = lstErrores.SelectedItem.SubItems(3)
    txtdatoMaestro = lstErrores.SelectedItem.SubItems(2)
End If

If Me.optActualiza(4).value = True Then  'CIIU
    txtDato = lstErrores.SelectedItem.SubItems(11)
    txtdatoCambio = lstErrores.SelectedItem.SubItems(12)
    txtdatoBase = lstErrores.SelectedItem.SubItems(12)
    txtdatoMaestro = lstErrores.SelectedItem.SubItems(18)
End If

End Sub

Private Sub lstErrores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtdatoCambio.SetFocus
End If
End Sub

Private Sub optActualiza_Click(Index As Integer)
    Me.txtdatoMaestro = ""
    If lnRegistros = 0 Then
        Exit Sub
    End If
    Select Case Index
        Case 0
            txtDato = lstErrores.SelectedItem.SubItems(4)
            txtdatoCambio = lstErrores.SelectedItem.SubItems(5)
            txtdatoBase = lstErrores.SelectedItem.SubItems(6)
            Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(15)
            txtcodsbs = lstErrores.SelectedItem.SubItems(2)
            If lstErrores.SelectedItem.SubItems(3) = "" Then
                txtcodsbsbase = lstErrores.SelectedItem.SubItems(2)
            Else
                txtcodsbsbase = lstErrores.SelectedItem.SubItems(3)
            End If
            Me.txtdatoMaestro.Locked = False
        Case 1  ' Documento
            txtDato = lstErrores.SelectedItem.SubItems(7)
            txtdatoCambio = lstErrores.SelectedItem.SubItems(8)
            txtdatoBase = lstErrores.SelectedItem.SubItems(8)
            Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(16)
            Me.txtdatoMaestro.Locked = True
        Case 2   ' RRPP
            txtDato = lstErrores.SelectedItem.SubItems(9)
            txtdatoCambio = lstErrores.SelectedItem.SubItems(10)
            txtdatoBase = lstErrores.SelectedItem.SubItems(10)
            Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(17)
            Me.txtdatoMaestro.Locked = True
        Case 4  ' Cod CIUU
            txtDato = lstErrores.SelectedItem.SubItems(11)
            txtdatoCambio = lstErrores.SelectedItem.SubItems(12)
            txtdatoBase = lstErrores.SelectedItem.SubItems(12)
            Me.txtdatoMaestro = lstErrores.SelectedItem.SubItems(18)
            Me.txtdatoMaestro.Locked = True
        Case 3 ' SBS
            txtDato = lstErrores.SelectedItem.SubItems(2)
            txtdatoCambio = lstErrores.SelectedItem.SubItems(3)
            txtdatoBase = lstErrores.SelectedItem.SubItems(3)
            txtdatoMaestro = lstErrores.SelectedItem.SubItems(2)
            Me.txtdatoMaestro.Locked = True
    End Select
End Sub

Private Sub optActualiza_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdatoBase.SetFocus
End If
End Sub

Private Sub txtCodSBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtdatoBase.SetFocus
End If
End Sub

Private Sub txtcodsbsbase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCambiar.SetFocus
End If
End Sub

Private Sub txtdatoBase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcodsbsbase.SetFocus
End If
End Sub

Private Sub txtdatoCambio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtdatoBase.SetFocus
End If
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txthasta.SetFocus
End If
End Sub

Private Sub txtDesde_LostFocus()
txtDesde = Format$(txtDesde, "000000")
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdProcesar.SetFocus
End If

End Sub

Private Sub txthasta_LostFocus()
txthasta = Format$(txthasta, "000000")
End Sub
