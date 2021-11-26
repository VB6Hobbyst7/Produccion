VERSION 5.00
Begin VB.Form frmUbicacionGeo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Ubicacion Geografica"
   ClientHeight    =   1845
   ClientLeft      =   2910
   ClientTop       =   3555
   ClientWidth     =   5715
   Icon            =   "frmUbicacionGeo.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   2010
      TabIndex        =   11
      Top             =   1350
      Width           =   1650
   End
   Begin VB.Frame fraZonaCbo 
      Height          =   1155
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   5610
      Begin VB.Frame frazona 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   210
         TabIndex        =   1
         Top             =   225
         Width           =   5325
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Zona"
            Top             =   75
            Width           =   1920
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Provincia"
            Top             =   465
            Width           =   1935
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Distrito"
            Top             =   105
            Width           =   1980
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   3
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Urbanización"
            Top             =   450
            Width           =   1995
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dpto :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   15
            TabIndex        =   9
            Top             =   150
            Width           =   375
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Prov :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2625
            TabIndex        =   8
            Top             =   165
            Width           =   375
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   -15
            TabIndex        =   7
            Top             =   495
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Zona :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   165
            Left            =   2625
            TabIndex        =   6
            Top             =   480
            Width           =   390
         End
      End
      Begin VB.Label Label1 
         Caption         =   " Zona :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   10
         Top             =   30
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmUbicacionGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TGarantiaTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Dim Nivel1() As String
Dim ContNiv1 As Integer
Dim Nivel2() As String
Dim ContNiv2 As Integer
Dim Nivel3() As String
Dim ContNiv3 As Integer
Dim Nivel4() As String
Dim ContNiv4 As Integer
Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer
Dim vsUbiGeoCod As String
Private sCodUbicacion As String

Public Function Inicio(Optional ByVal psubiGeoCod As String = "") As String
    vsUbiGeoCod = psubiGeoCod
    Me.Show 1
    Inicio = sCodUbicacion
End Function

Private Sub CargaZonas(Optional ByVal psubiGeoCod As String = "")
Dim i As Integer
Dim nEnc As Integer
    If Trim(psubiGeoCod) = "" Then
        Exit Sub
    End If
    
    'Ubica Dpto
    nEnc = -1
    For i = 0 To cmbPersUbiGeo(0).ListCount - 1
        If Right(cmbPersUbiGeo(0).List(i), 12) = "1" & Mid(psubiGeoCod, 2, 2) & String(9, "0") Then
            nEnc = i
            Exit For
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nEnc
    'Ubica Prov
    nEnc = -1
    For i = 0 To cmbPersUbiGeo(1).ListCount - 1
        If Right(cmbPersUbiGeo(1).List(i), 12) = "2" & Mid(psubiGeoCod, 2, 4) & String(7, "0") Then
            nEnc = i
            Exit For
        End If
    Next i
    cmbPersUbiGeo(1).ListIndex = nEnc
    
    'Ubica Distrito
    nEnc = -1
    For i = 0 To cmbPersUbiGeo(2).ListCount - 1
        If Right(cmbPersUbiGeo(2).List(i), 12) = "3" & Mid(psubiGeoCod, 2, 6) & String(5, "0") Then
            nEnc = i
            Exit For
        End If
    Next i
    cmbPersUbiGeo(2).ListIndex = nEnc
    'Ubica Zona
    nEnc = -1
    For i = 0 To cmbPersUbiGeo(3).ListCount - 1
        If Right(cmbPersUbiGeo(3).List(i), 12) = psubiGeoCod Then
            nEnc = i
            Exit For
        End If
    Next i
    cmbPersUbiGeo(3).ListIndex = nEnc
    
    
End Sub

Private Sub CargaUbicacionesGeograficas()
Dim Conn As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim nPos As Integer

On Error GoTo ErrCargaUbicacionesGeograficas
    Set Conn = New DConecta
    'Carga Niveles
    sSQL = "Select *, 1 p from UbicacionGeografica where cUbiGeoCod like '1%'"
    sSQL = sSQL & " Union "
    sSQL = sSQL & " select *, 2 p from UbicacionGeografica where cUbiGeoCod like '2%' "
    sSQL = sSQL & " Union "
    sSQL = sSQL & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '3%' "
    sSQL = sSQL & " Union "
    sSQL = sSQL & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    ContNiv4 = 0
    
    Conn.AbreConexion
    Set R = Conn.CargaRecordSet(sSQL)
    Do While Not R.EOF
        Select Case R!P
            Case 1 ' Departamento
                ContNiv1 = ContNiv1 + 1
                ReDim Preserve Nivel1(ContNiv1)
                Nivel1(ContNiv1 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 2 ' Provincia
                ContNiv2 = ContNiv2 + 1
                ReDim Preserve Nivel2(ContNiv2)
                Nivel2(ContNiv2 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 3 'Distrito
                ContNiv3 = ContNiv3 + 1
                ReDim Preserve Nivel3(ContNiv3)
                Nivel3(ContNiv3 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 4 'Zona
                ContNiv4 = ContNiv4 + 1
                ReDim Preserve Nivel4(ContNiv4)
                Nivel4(ContNiv4 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
        End Select
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For i = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(i)
        If Trim(Right(Nivel1(i), 12)) = "113000000000" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    Exit Sub
    
ErrCargaUbicacionesGeograficas:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub

Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TGarantiaTipoCombo)
Dim i As Integer
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboProv
            cmbPersUbiGeo(1).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(1).Clear
                For i = 0 To ContNiv2 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel2(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(1).AddItem Nivel2(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel3(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel4(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
    End Select
End Sub


Private Sub cmdAceptar_Click()
    sCodUbicacion = cmbPersUbiGeo(3).Text
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargaUbicacionesGeograficas
    Call CargaZonas(vsUbiGeoCod)
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
        Select Case Index
            Case 0 'Combo Dpto
                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboProv)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(2).Clear
                    cmbPersUbiGeo(3).Clear
                End If
            Case 1 'Combo Provincia
                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboDist)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(3).Clear
                End If
            Case 2 'Combo Distrito
                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboZona)
        End Select
End Sub


Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        Select Case Index
            Case 0
                cmbPersUbiGeo(1).SetFocus
            Case 1
                cmbPersUbiGeo(2).SetFocus
            Case 2
                cmbPersUbiGeo(3).SetFocus
            'Case 3
            '    txtMontotas.SetFocus
        End Select
     End If
End Sub

