VERSION 5.00
Begin VB.Form frmPersActualizaDireccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Dirección"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   Icon            =   "frmPersActualizaDireccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   300
      Left            =   4800
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Información Domicilio "
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   80
      TabIndex        =   5
      Top             =   1200
      Width           =   5685
      Begin VB.ComboBox cmbPersUbiGeo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   3
         ItemData        =   "frmPersActualizaDireccion.frx":030A
         Left            =   3120
         List            =   "frmPersActualizaDireccion.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   2
         ItemData        =   "frmPersActualizaDireccion.frx":030E
         Left            =   120
         List            =   "frmPersActualizaDireccion.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   4
         ItemData        =   "frmPersActualizaDireccion.frx":0312
         Left            =   120
         List            =   "frmPersActualizaDireccion.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         ItemData        =   "frmPersActualizaDireccion.frx":0316
         Left            =   3120
         List            =   "frmPersActualizaDireccion.frx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   2415
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   0
         ItemData        =   "frmPersActualizaDireccion.frx":031A
         Left            =   120
         List            =   "frmPersActualizaDireccion.frx":031C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox txtPersDireccDomicilio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2640
         Width           =   5355
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Distrito :"
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Zona : "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pais : "
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Departamento :"
         Height          =   195
         Left            =   3120
         TabIndex        =   14
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Provincia :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblPersDireccDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   630
      End
   End
   Begin VB.Frame frmPersActualizaDireccion 
      Caption         =   "Datos Clientes:"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Label lblPersNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   280
         Left            =   960
         TabIndex        =   4
         Top             =   690
         Width           =   4575
      End
      Begin VB.Label lblPersCodigo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   280
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPersActualizaDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oPersona As UPersona_Cli
Dim sPersDirec As String
Dim sCtaCod As String
Dim nActualiza As Integer

Dim sUbigeo As String
Dim sDireccion As String

Public Function IniciaFormulario(ByVal psPersCod As String, ByVal psCtaCod As String) As String
    Call CargarDatos
    sCtaCod = psCtaCod
    If Cargar_Datos_Persona(psPersCod) Then
        Me.Show 1
    Else
        MsgBox "No se encontró datos de la persona", vbInformation, "Alerta"
    End If
    IniciaFormulario = sPersDirec
    sPersDirec = ""
End Function

Private Sub cmbPersUbiGeo_Click(Index As Integer)
    Dim oUbic As comdpersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Index <> 4 Then
        Set oUbic = New comdpersona.DCOMPersonas
        Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

        If Trim(Right(cmbPersUbiGeo(0).Text, 12)) <> "04028" Then
            If Index = 0 Then
                For i = 1 To cmbPersUbiGeo.Count - 1
                    cmbPersUbiGeo(i).Clear
                    cmbPersUbiGeo(i).AddItem Trim(Trim(cmbPersUbiGeo(0).Text)) & Space(50) & Trim(Right(cmbPersUbiGeo(0).Text, 12))
                Next i
            End If
        Else
            For i = Index + 1 To cmbPersUbiGeo.Count - 1
                cmbPersUbiGeo(i).Clear
            Next
            While Not rs.EOF
                cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
                rs.MoveNext
            Wend
        End If
        Set oUbic = Nothing
    End If
End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            txtPersDireccDomicilio.SetFocus
        End If
    End If
End Sub

Private Sub Limpiar()
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), "04028")
    cmbPersUbiGeo(1).ListIndex = -1
    cmbPersUbiGeo(2).ListIndex = -1
    cmbPersUbiGeo(3).ListIndex = -1
    cmbPersUbiGeo(4).ListIndex = -1
    
    cmbPersUbiGeo(0).Enabled = True
    cmbPersUbiGeo(1).Enabled = True
    cmbPersUbiGeo(2).Enabled = True
    cmbPersUbiGeo(3).Enabled = True
    cmbPersUbiGeo(4).Enabled = True
    txtPersDireccDomicilio.Enabled = False
End Sub
    
Private Sub CargarDatos()
    Dim oDPersona As New comdpersona.DCOMPersonas
    Dim lrsUbiGeo  As ADODB.Recordset
    Dim nPos As Integer
    
    Set lrsUbiGeo = oDPersona.CargarUbicacionesGeograficas(True, 0)
    
    While Not lrsUbiGeo.EOF
        cmbPersUbiGeo(0).AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        lrsUbiGeo.MoveNext
    Wend
    
    lrsUbiGeo.MoveFirst

    If lrsUbiGeo.RecordCount > 0 Then lrsUbiGeo.MoveFirst
    For i = 0 To lrsUbiGeo.RecordCount
        If Trim(lrsUbiGeo!cUbiGeoCod) = "04028" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
End Sub

Function Cargar_Datos_Persona(pcPersCod As String) As Boolean
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli
    oPersona.sCodAge = gsCodAge
    
    Cargar_Datos_Persona = True
    Call oPersona.RecuperaPersona(pcPersCod, , gsCodUser)
    If oPersona.PersCodigo = "" Then
        Cargar_Datos_Persona = False
        Exit Function
    End If
    
     If Len(Trim(oPersona.UbicacionGeografica)) = 12 Then
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "04028")
        cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbicacionGeografica, 2, 2) & String(9, "0"))
        cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbicacionGeografica, 2, 4) & String(7, "0"))
        cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbicacionGeografica, 2, 6) & String(5, "0"))
        cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & oPersona.UbicacionGeografica)
    Else
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & oPersona.UbicacionGeografica)
        cmbPersUbiGeo(1).Clear
        cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(1).ListIndex = 0
        cmbPersUbiGeo(2).Clear
        cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(2).ListIndex = 0
        cmbPersUbiGeo(3).Clear
        cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(3).ListIndex = 0
        cmbPersUbiGeo(4).Clear
        cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(4).ListIndex = 0
    End If
    lblPersCodigo.Caption = oPersona.PersCodigo
    lblPersNombre.Caption = oPersona.NombreCompleto
    txtPersDireccDomicilio.Text = oPersona.Domicilio
    txtPersDireccDomicilio.Enabled = True
    
    sUbigeo = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    sDireccion = oPersona.Domicilio
End Function

Private Sub cmdGuardar_Click()
    Dim oDPersona As New comdpersona.DCOMPersonas
    Dim oColP As New COMNColoCPig.NCOMColPValida
    Dim sMsj As String
    
    sMsj = ValidaDatos
    
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    
    If txtPersDireccDomicilio.Text <> "" Then
        'Call ValidaCambios
        nActualiza = 1
        
        If MsgBox("¿Está seguro de grabar?", vbQuestion + vbYesNo, "Alerta") = vbNo Then
             Exit Sub
        End If
        
        If sUbigeo = Trim(Right(cmbPersUbiGeo(4).Text, 15)) And UCase(sDireccion) = UCase(txtPersDireccDomicilio.Text) Then
            If MsgBox("¿Está seguro que desea grabar sin actualizar los datos?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
                nActualiza = 2
            Else
                Exit Sub
            End If
        End If
        
        Call oDPersona.ActualizaUBIGEOPersona(Trim(Right(cmbPersUbiGeo(4).Text, 15)), txtPersDireccDomicilio.Text, lblPersCodigo.Caption)
        MsgBox "Los datos se guardaron de forma exitosa.", vbInformation, "Alerta"
        sPersDirec = txtPersDireccDomicilio.Text
        Call oColP.PignoActualizaDirecCliente(lblPersCodigo.Caption, sCtaCod, nActualiza)
        Call Limpiar
        Unload Me
    Else
        MsgBox "Los datos no pueden estar vacíos", vbInformation, "Alerta"
    End If
End Sub

Private Sub Form_Deactivate()
    If MsgBox("hola", vbYesNo) = vbNo Then
        Exit Sub
     End If
End Sub

Public Function ValidaDatos() As String
    ValidaDatos = ""
    If cmbPersUbiGeo(0).Text = "" Or cmbPersUbiGeo(1).Text = "" Or cmbPersUbiGeo(2).Text = "" Or cmbPersUbiGeo(3).Text = "" Or cmbPersUbiGeo(4).Text = "" Then
        ValidaDatos = "No se puede guardar datos vacíos"
        Exit Function
    End If
End Function

Public Sub ValidaCambios()
    nActualiza = 1
    
    If sUbigeo = Trim(Right(cmbPersUbiGeo(4).Text, 15)) And sDireccion = oPersona.Domicilio Then
        If MsgBox("¿Está seguro que desea grabar sin actualizar los datos?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
            nActualiza = 2
        End If
    End If
End Sub
