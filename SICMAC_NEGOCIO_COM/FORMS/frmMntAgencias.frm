VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntAgencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Agencias"
   ClientHeight    =   5280
   ClientLeft      =   1080
   ClientTop       =   1575
   ClientWidth     =   10020
   Icon            =   "frmMntAgencias.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   75
      TabIndex        =   2
      Top             =   4500
      Width           =   9885
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   8445
         TabIndex        =   8
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   8445
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   7140
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   405
         Left            =   2850
         TabIndex        =   5
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   405
         Left            =   1530
         TabIndex        =   4
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   405
         Left            =   210
         TabIndex        =   3
         Top             =   225
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agencias"
      Height          =   4395
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   9900
      Begin MSDataGridLib.DataGrid DGAgencias 
         Height          =   4050
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   7144
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cAgeCod"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cAgeDescripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cAgeDireccion"
            Caption         =   "Direccion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cAgeTelefono"
            Caption         =   "Telefono"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cSubCtaCod"
            Caption         =   "SubCta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nAgeEspecial"
            Caption         =   "Age Esp"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3839.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3135.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1785.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   824.882
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtCod 
         Height          =   285
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox TxtDescrip 
         Height          =   285
         Left            =   1380
         TabIndex        =   12
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox TxtDirecc 
         Height          =   285
         Left            =   5970
         TabIndex        =   14
         Top             =   2160
         Width           =   3570
      End
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   1380
         TabIndex        =   16
         Top             =   2535
         Width           =   3495
      End
      Begin VB.TextBox TxtSubCta 
         Height          =   285
         Left            =   5970
         TabIndex        =   18
         Top             =   2565
         Width           =   795
      End
      Begin VB.CheckBox ChkAgeEspe 
         Alignment       =   1  'Right Justify
         Caption         =   "Agencia Especial"
         Height          =   210
         Left            =   6885
         TabIndex        =   19
         Top             =   2595
         Width           =   1590
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   0
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3150
         Width           =   2175
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   3
         Left            =   3765
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3750
         Width           =   2430
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   1
         Left            =   3765
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3135
         Width           =   2430
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   2
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3735
         Width           =   2190
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "U. Geografica :"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   2910
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Distrito :"
         Height          =   195
         Left            =   1290
         TabIndex        =   27
         Top             =   3495
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Zona : "
         Height          =   195
         Left            =   3660
         TabIndex        =   26
         Top             =   3510
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Departamento :"
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Provincia :"
         Height          =   195
         Left            =   3675
         TabIndex        =   24
         Top             =   2880
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sub Cta :"
         Height          =   195
         Left            =   5145
         TabIndex        =   17
         Top             =   2580
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telefonos :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2550
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Direccion :"
         Height          =   195
         Left            =   5100
         TabIndex        =   13
         Top             =   2175
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion :"
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   2175
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   1800
         Width           =   645
      End
      Begin VB.Shape Shape1 
         Height          =   2580
         Left            =   120
         Top             =   1680
         Width           =   9570
      End
   End
End
Attribute VB_Name = "frmMntAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum
Dim RAge As ADODB.Recordset
Dim nAccion As Integer
Dim Nivel1() As String
Dim ContNiv1 As Integer
Dim Nivel2() As String
Dim ContNiv2 As Integer
Dim Nivel3() As String
Dim ContNiv3 As Integer
Dim Nivel4() As String
Dim ContNiv4 As Integer
Dim Nivel5() As String
Dim ContNiv5 As Integer
Dim bEstadoCargando As Boolean

Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TTipoCombo)
Dim i As Integer
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboDpto
            cmbPersUbiGeo(0).Clear
            If sCodigo = "PER" Then
                For i = 0 To ContNiv1 - 1
                    cmbPersUbiGeo(0).AddItem Nivel1(i)
                Next i
            Else
                cmbPersUbiGeo(0).AddItem psValor
            End If
        Case ComboProv
            cmbPersUbiGeo(1).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(1).Clear
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel3(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(1).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel4(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(2).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv5 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel5(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(3).AddItem Nivel5(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
    End Select
End Sub


'Private Sub CargaUbicacionesGeograficas()
'Dim Conn As DConecta
'Dim sSQL As String
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim nPos As Integer
'
'On Error GoTo ErrCargaUbicacionesGeograficas
'    Set Conn = New DConecta
'    'Carga Niveles
'    sSQL = "select *,1 p from UbicacionGeografica where len(cUbiGeoCod)=3 "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
'    sSQL = sSQL & " Union "
'    sSQL = sSQL & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
'    ContNiv1 = 0
'    ContNiv2 = 0
'    ContNiv3 = 0
'    ContNiv4 = 0
'    ContNiv5 = 0
'
'    Conn.AbreConexion
'    Set R = Conn.CargaRecordSet(sSQL)
'    Do While Not R.EOF
'        Select Case R!P
'            Case 1 'Pais
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 2 ' Departamento
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 3 'Provincia
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 4 'Distrito
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 5 'Zona
'                ContNiv5 = ContNiv5 + 1
'                ReDim Preserve Nivel5(ContNiv5)
'                Nivel5(ContNiv5 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'        End Select
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Conn.CierraConexion
'    Set Conn = Nothing
'
'    'Carga el Nivel1 en el Control
'    cmbPersUbiGeo(0).Clear
'    For i = 0 To ContNiv2 - 1
'        cmbPersUbiGeo(0).AddItem Nivel2(i)
'        If Mid(Trim(Right(Nivel2(i), 15)), 1, 3) = "113" Then
'            nPos = i
'        End If
'    Next i
'    cmbPersUbiGeo(0).ListIndex = nPos
'    cmbPersUbiGeo(1).Clear
'    cmbPersUbiGeo(2).Clear
'    cmbPersUbiGeo(3).Clear
'    Exit Sub
'
'ErrCargaUbicacionesGeograficas:
'    MsgBox Err.Description, vbInformation, "Aviso"
'
'End Sub

Private Sub HabilitaDatos(ByVal pbHabilita As Boolean)
    If pbHabilita Then
        DGAgencias.Height = 1335
    Else
        DGAgencias.Height = 4050
    End If
    cmdNuevo.Visible = Not pbHabilita
    cmdeditar.Visible = Not pbHabilita
    cmdeliminar.Visible = Not pbHabilita
    cmdSalir.Visible = Not pbHabilita
    CmdAceptar.Visible = pbHabilita
    cmdcancelar.Visible = pbHabilita
End Sub

Private Sub LimpiaDatos()
    TxtCod.Text = ""
    TxtDescrip.Text = ""
    TxtDirecc.Text = ""
    TxtSubCta.Text = ""
    TxtTelefono.Text = ""
    ChkAgeEspe.value = 0
    InicializaCombos Me
End Sub

Private Sub CargaGrid()
Dim oDAgencias As COMDConstantes.DCOMAgencias

    Set oDAgencias = New COMDConstantes.DCOMAgencias
    Set RAge = oDAgencias.RecuperaAgencias
    Set oDAgencias = Nothing
    Set DGAgencias.DataSource = RAge
    
End Sub

Private Sub CmdAceptar_Click()
Dim oBase As COMDCredito.DCOMCredActBD
Dim oDGeneral As COMDConstSistema.DCOMGeneral
Dim sMovAct As String
    Set oDGeneral = New COMDConstSistema.DCOMGeneral
    sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, Mid(gsCodAge, 1, 3), gsCodAge)
    Set oDGeneral = Nothing
    Set oBase = New COMDCredito.DCOMCredActBD
    If nAccion = 1 Then
        Call oBase.dInsertAgencia(Trim(TxtCod.Text), Trim(TxtDescrip.Text), Trim(TxtDirecc.Text), _
            Trim(TxtTelefono.Text), Trim(TxtSubCta.Text), ChkAgeEspe.value, cmbPersUbiGeo(3), sMovAct, False)
    Else
        Call oBase.dUpdateAgencia(Trim(TxtCod.Text), Trim(TxtDescrip.Text), Trim(TxtDirecc.Text), _
            Trim(TxtTelefono.Text), Trim(TxtSubCta.Text), ChkAgeEspe.value, cmbPersUbiGeo(3), sMovAct, False)
    End If
    Set oBase = Nothing
    Call CargaGrid
    HabilitaDatos False
End Sub

Private Sub CmdCancelar_Click()
    HabilitaDatos False
End Sub

Private Sub CmdEditar_Click()
    HabilitaDatos True
    TxtCod.Enabled = False
    TxtCod.Text = RAge!cAgeCod
    TxtDescrip.Text = RAge!cAgeDescripcion
    TxtDirecc.Text = RAge!cAgeDireccion
    TxtTelefono.Text = RAge!cAgeTelefono
    TxtSubCta.Text = RAge!cSubCtaCod
    ChkAgeEspe.value = IIf(RAge!nAgeEspecial, 1, 0)
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "1" & Mid(RAge!cUbiGeoCod, 2, 2) & String(9, "0"))
    cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "2" & Mid(RAge!cUbiGeoCod, 2, 4) & String(7, "0"))
    cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "3" & Mid(RAge!cUbiGeoCod, 2, 6) & String(5, "0"))
    cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & RAge!cUbiGeoCod)
    TxtDescrip.SetFocus
    nAccion = 2
End Sub

Private Sub cmdeliminar_Click()
Dim oBase As COMDCredito.DCOMCredActBD
    DGAgencias.Col = 1
    If MsgBox("Se va Ha Eliminar la Agencia " & DGAgencias.Text & ", Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oBase = New COMDCredito.DCOMCredActBD
    DGAgencias.Col = 0
    Call oBase.dDeleteAgencia(Trim(DGAgencias.Text), False)
    Set oBase = Nothing
    Call CargaGrid
End Sub

Private Sub cmdNuevo_Click()
    HabilitaDatos True
    Call LimpiaDatos
    TxtCod.Enabled = True
    TxtCod.SetFocus
    nAccion = 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oUbi As COMDPersona.DCOMPersonas
Dim rsUbiGeo As ADODB.Recordset
    
    Call CargaGrid
    'Call CargaUbicacionesGeograficas
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    Set oUbi = New COMDPersona.DCOMPersonas
    Set rsUbiGeo = oUbi.CargarUbicacionesGeograficas(, 1)
    Set oUbi = Nothing
    
    While Not rsUbiGeo.EOF
        cmbPersUbiGeo(0).AddItem Trim(rsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(rsUbiGeo!cUbiGeoCod)
        rsUbiGeo.MoveNext
    Wend

End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
'        Select Case Index
'            Case 0 'Combo Dpto
'                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboProv)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(2).Clear
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 1 'Combo Prov
'                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboDist)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 2 'Combo Distrito
'                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboZona)
'        End Select
Dim oUbic As COMDPersona.DCOMPersonas
Dim Rs As ADODB.Recordset
Dim i As Integer

If Index = 3 Then Exit Sub

Set oUbic = New COMDPersona.DCOMPersonas

Set Rs = oUbic.CargarUbicacionesGeograficas(, Index + 2, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

For i = Index + 1 To cmbPersUbiGeo.Count - 1
    cmbPersUbiGeo(i).Clear
Next

While Not Rs.EOF
    cmbPersUbiGeo(Index + 1).AddItem Trim(Rs!cUbiGeoDescripcion) & Space(50) & Trim(Rs!cUbiGeoCod)
    Rs.MoveNext
Wend

Set oUbic = Nothing

End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 3 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            CmdAceptar.SetFocus
        End If
    End If
End Sub

Private Sub TxtCod_GotFocus()
    fEnfoque TxtCod
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtDescrip.SetFocus
    End If
End Sub

Private Sub TxtDescrip_GotFocus()
    fEnfoque TxtDescrip
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtDirecc.SetFocus
    End If
End Sub


Private Sub TxtDirecc_GotFocus()
    fEnfoque TxtDirecc
End Sub

Private Sub TxtDirecc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtTelefono.SetFocus
    End If
End Sub

Private Sub TxtSubCta_GotFocus()
    fEnfoque TxtSubCta
End Sub

Private Sub TxtSubCta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersUbiGeo(0).SetFocus
    End If
End Sub

Private Sub TxtTelefono_GotFocus()
    fEnfoque TxtTelefono
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtSubCta.SetFocus
    End If
End Sub
