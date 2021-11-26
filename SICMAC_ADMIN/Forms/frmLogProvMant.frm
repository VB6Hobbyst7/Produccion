VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProvMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   5910
   ClientLeft      =   735
   ClientTop       =   1200
   ClientWidth     =   10590
   Icon            =   "frmLogProvMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   330
      Left            =   90
      TabIndex        =   15
      Top             =   5505
      Width           =   1035
   End
   Begin VB.Frame fraContenedor 
      Appearance      =   0  'Flat
      Caption         =   "Proveedores "
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
      Height          =   5370
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Width           =   10440
      Begin VB.CommandButton cmdProv 
         Caption         =   "Buen Cont."
         Height          =   345
         Index           =   7
         Left            =   7380
         TabIndex        =   14
         Top             =   2850
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Age.Reten."
         Height          =   345
         Index           =   6
         Left            =   6345
         TabIndex        =   13
         Top             =   2850
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Cuentas"
         Height          =   345
         Index           =   5
         Left            =   5310
         TabIndex        =   12
         Top             =   2850
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   2850
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Act./Des."
         Height          =   345
         Index           =   1
         Left            =   1170
         TabIndex        =   7
         Top             =   2850
         Width           =   990
      End
      Begin VB.CommandButton cmdProvBS 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   0
         Left            =   9330
         TabIndex        =   6
         Top             =   3690
         Width           =   930
      End
      Begin VB.CommandButton cmdProvBS 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   1
         Left            =   9330
         TabIndex        =   5
         Top             =   4065
         Width           =   930
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   2
         Left            =   2205
         TabIndex        =   4
         Top             =   2850
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Buscar"
         Height          =   345
         Index           =   3
         Left            =   3240
         TabIndex        =   3
         Top             =   2850
         Width           =   990
      End
      Begin VB.CommandButton cmdProv 
         Caption         =   "Imprimir"
         Height          =   345
         Index           =   4
         Left            =   4275
         TabIndex        =   2
         Top             =   2850
         Width           =   990
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProv 
         Height          =   2580
         Left            =   120
         TabIndex        =   9
         Top             =   225
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   4551
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         SelectionMode   =   1
         AllowUserResizing=   1
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProvBS 
         Height          =   1590
         Left            =   120
         TabIndex        =   10
         Top             =   3690
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   2805
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Bien/Servicio del Proveedor"
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
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   3435
         Width           =   2640
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   10230
         Y1              =   3345
         Y2              =   3345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   135
         X2              =   10200
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   9525
      TabIndex        =   0
      Top             =   5475
      Width           =   1035
   End
End
Attribute VB_Name = "frmLogProvMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDProv As DLogProveedor
Dim clsNProv As NLogProveedor
'ARLO 20170125******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdActualizar_Click()
    clsDProv.ActualizaBSProveedores GetMovNro(gsCodUser, gsCodAge)
    MsgBox "Actualiacion satisfactoria.", vbInformation, "Aviso"
End Sub

Private Sub cmdProv_Click(Index As Integer)
    Dim rsProvBS As ADODB.Recordset
    Dim obp As UPersona
    Dim nCont As Integer, nResult As Integer
    Dim sactualiza As String
    Dim sPersCod As String
    Dim sPersNombre As String
    Dim nProvEstado As Integer
    Dim nProvAgeReten As LogProvAgenteRetencion
    Dim nProvBuenCont As LogProvBuenContribuyente
    Dim bPersDoc As Boolean
    Dim rsProv As ADODB.Recordset
    
    If fgProv.TextMatrix(fgProv.row, 0) = "" And Index <> 0 Then Exit Sub
    
    If Index > 0 Then
        If fgProv.row = 0 Then
            MsgBox "No existen proveedores", vbInformation, " Aviso"
            Exit Sub
        End If
        sPersCod = fgProv.TextMatrix(fgProv.row, 0)
        sPersNombre = fgProv.TextMatrix(fgProv.row, 1)
        'INICIO ORCR-20140913*********
        'nProvEstado = Val(fgProv.TextMatrix(fgProv.Row, 5))
        'nProvAgeReten = IIf(fgProv.TextMatrix(fgProv.Row, 9), gLogProvAgenteRetencionSi, gLogProvAgenteRetencionNo)
        nProvEstado = Val(fgProv.TextMatrix(fgProv.row, 6))
        'nProvAgeReten = IIf(fgProv.TextMatrix(fgProv.row, 10), gLogProvAgenteRetencionSi, gLogProvAgenteRetencionNo)  'comentado xPASIERS0472015
        nProvAgeReten = IIf(fgProv.TextMatrix(fgProv.row, 13), gLogProvAgenteRetencionSi, gLogProvAgenteRetencionNo) 'PASIERS0472015
        'FIN ORCR-20140913************
        'nProvBuenCont = IIf(fgProv.TextMatrix(fgProv.Row, 11), gLogProvBuenContribuyenteSi, gLogProvBuenContribuyenteNo)
        
    End If

    Select Case Index
        Case 0:
            'Agregar proveedores
            Set obp = frmBuscaPersona.Inicio
            If obp Is Nothing Then Exit Sub
            sPersCod = Trim(obp.sPersCod)
            For nCont = 1 To fgProv.Rows - 1
                If fgProv.TextMatrix(nCont, 0) = sPersCod Then
                    MsgBox "Proveedor " & obp.sPersNombre & " ya se encuentra registrado", vbInformation, " Aviso "
                    Exit Sub
                End If
            Next
            
            bPersDoc = True
            If Trim(obp.sPersIdnroRUC) = "" Then
'INICIO ORCR-20140913*********
                bPersDoc = MsgBox("Esta Persona no cuenta con RUC, ¿Está Seguro de Registrarlo como Proveedor?", vbInformation + vbYesNo, "Aviso") = vbYes
'                bPersDoc = False
'                For nCont = 1 To obp.DocsPers.RecordCount
'                    If obp.DocsPers!cPersIdTpo = gPersIdRUS Then
'                        bPersDoc = True
'                        Exit For
'                    End If
'                    obp.DocsPers.MoveNext
'                Next
'FIN ORCR-20140913************
            End If



            If sPersCod <> "" And bPersDoc Then
                sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nResult = clsDProv.GrabaProveedor(sPersCod, sactualiza)
                If nResult = 0 Then
                    Set fgProv.Recordset = clsDProv.CargaProveedor
                        'ARLO 20170125
                        gsopecod = LogPistaManProveedores
                        Set objPista = New COMManejador.Pista
                        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Agrego al Proveedor Cod Persona : " & sPersCod, sPersCod, 3
                        Set objPista = Nothing
                        '***********
                    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
                    
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Rows = 2
                        fgProvBS.Clear
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            Else
                MsgBox "No se hizo la agregación", vbInformation, "Aviso"
            End If
        Case 1:
            'Activar/Desactivar
            nResult = clsNProv.ActDesProveedor(sPersCod, nProvEstado)
            'ARLO 20170125
            gsopecod = LogPistaManProveedores
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", "Modifico al Proveedor Cod Persona : " & sPersCod, sPersCod, 3
            Set objPista = Nothing
            '***********
            If nResult = 0 Then
                Set fgProv.Recordset = clsDProv.CargaProveedor
            Else
                Call ErrorDesc(nResult)
            End If
        Case 2:
            'Eliminar
            '***************************** N O T A *****************************
            'Validar que no halla realizado ningun proceso
            If MsgBox("¿ Estás seguro de eliminar a " & sPersNombre & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDProv.EliminaProveedor(sPersCod)
                'ARLO 20170125
                gsopecod = LogPistaManProveedores
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Elimino al Proveedor Cod Persona : " & sPersCod, sPersCod, 3
                Set objPista = Nothing
                '***********
                If nResult = 0 Then
                    Set rsProv = New ADODB.Recordset
                    Set rsProv = clsDProv.CargaProveedor

                    If rsProv.EOF And rsProv.BOF Then
                        fgProv.Rows = 1
                        fgProv.Rows = 2
                        fgProv.FixedRows = 1
                    Else
                        Set fgProv.Recordset = clsDProv.CargaProveedor
                    End If
                    
                    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
                    
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Clear
                        fgProvBS.Rows = 2
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            End If
        Case 3:
            'BUSCAR
            sPersNombre = Trim(frmLogMantOpc.Inicio(0, 1))
            If sPersNombre <> "" Then
                'ARLO 20170125
                gsopecod = LogPistaManProveedores
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Busco al Producto " & sPersNombre, ""
                Set objPista = Nothing
                '***********
                For nCont = 1 To fgProv.Rows - 1
                    If Left(fgProv.TextMatrix(nCont, 1), Len(sPersNombre)) = sPersNombre Then
                        'fgProv.Row = nCont
                        fgProv.TopRow = nCont
                        Exit Sub
                    End If
                Next
                MsgBox "Nombre no se encuentra en la lista", vbInformation, " Aviso"
            End If
        Case 4:
            'IMPRIMIR
            Dim clsNImp As NLogImpre
            Dim clsPrevio As clsPrevio
            Dim sImpre As String
            MousePointer = 11
            Set clsNImp = New NLogImpre
            sImpre = clsNImp.ImpProveedor(gsNomAge, gdFecSis)
            Set clsNImp = Nothing
            MousePointer = 0
            Set clsPrevio = New clsPrevio
            clsPrevio.Show sImpre, Me.Caption, True
            Set clsPrevio = Nothing
        
        Case 5:
            'ASIGNACION DE CTAS
            'INICIO ORCR-20140913*********
            'frmLogProvCtas.Ini fgProv.TextMatrix(fgProv.Row, 0), fgProv.TextMatrix(fgProv.Row, 1), fgProv.TextMatrix(fgProv.Row, 7), fgProv.TextMatrix(fgProv.Row, 8), fgProv.TextMatrix(fgProv.Row, 0)
            frmLogProvCtas.Ini fgProv.TextMatrix(fgProv.row, 0), fgProv.TextMatrix(fgProv.row, 1), fgProv.TextMatrix(fgProv.row, 8), fgProv.TextMatrix(fgProv.row, 9), fgProv.TextMatrix(fgProv.row, 0)
            'FIN ORCR-20140913************
            Set frmLogProvCtas = Nothing
            Set rsProv = clsDProv.CargaProveedor
            If rsProv.RecordCount > 0 Then
                Set fgProv.Recordset = rsProv
            Else
                cmdProv(0).Enabled = False
                cmdProv(1).Enabled = False
                cmdProv(2).Enabled = False
                cmdProv(3).Enabled = False
                cmdProv(4).Enabled = False
                cmdProvBS(0).Enabled = False
                cmdProvBS(1).Enabled = False
                Exit Sub
            End If
        Case 6:
            'Agente de Retencion
            nResult = clsNProv.ProveedorActAgeReten(sPersCod, nProvAgeReten)
            If nResult = 0 Then
                Set fgProv.Recordset = clsDProv.CargaProveedor
            Else
                Call ErrorDesc(nResult)
            End If
        Case 7:
            'Buen Contribuyente
            nResult = clsNProv.ProveedorBuenCont(sPersCod, nProvBuenCont)
            If nResult = 0 Then
                Set fgProv.Recordset = clsDProv.CargaProveedor
            Else
                Call ErrorDesc(nResult)
            End If
            
            
            
        Case Else
            MsgBox "Indice de comand de proveedores no reconocido", vbInformation, " Aviso "
    End Select
    
    MDISicmact.staMain.Panels(2).Text = fgProv.Rows & " registros."
End Sub

Private Sub cmdProvBS_Click(Index As Integer)
    Dim clsDBS As DLogBieSer
    Dim rsProvBS As ADODB.Recordset
    Dim nResult As Integer
    Dim sactualiza As String
    Dim sPersCod As String
    Dim sBSCod As String
    Dim sBSNombre As String
    
    If fgProv.TextMatrix(fgProv.row, 0) = "" Then Exit Sub

    
    If fgProv.row = 0 Then
        MsgBox "No existen proveedores", vbInformation, " Aviso"
        Exit Sub
    End If
    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
    
    If Index > 0 Then
        If fgProvBS.row = 0 Then
            MsgBox "No existen bienes/servicios del proveedor", vbInformation, " Aviso"
            Exit Sub
        End If
        sBSCod = fgProvBS.TextMatrix(fgProvBS.row, 0)
        sBSNombre = fgProvBS.TextMatrix(fgProvBS.row, 2)
    End If
    
    Select Case Index
        Case Is = 0
            'Agregar bien/servicio de proveedor
            Dim vBS As ClassDescObjeto
            Set vBS = New ClassDescObjeto
            Set clsDBS = New DLogBieSer
            vBS.ColCod = 0
            vBS.ColDesc = 1
            vBS.Show clsDBS.CargaBS(BsTodosArbol), ""
            
            Set clsDBS = Nothing
            If vBS.lbOK Then
                sBSCod = vBS.gsSelecCod
                sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                
                nResult = clsDProv.GrabaProveedorBS(sPersCod, sBSCod, sactualiza)
                If nResult = 0 Then
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Rows = 2
                        fgProvBS.Clear
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            Else
                MsgBox "No se hizo la agregación", vbInformation, "Aviso"
            End If
            Set vBS = Nothing
        Case Is = 1
            'Eliminar
            If sBSCod = "" Then Exit Sub

            If MsgBox("¿ Estás seguro de eliminar " & sBSNombre & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDProv.EliminaProveedorBS(sPersCod, sBSCod)
                If nResult = 0 Then
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Rows = 2
                        fgProvBS.Clear
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            End If
        Case Else
            MsgBox "Indice de comand bien/servicio de proveedor no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDProv = Nothing
    Set clsNProv = Nothing
    Unload Me
End Sub

Private Sub fgProv_Click()
    Dim rsProvBS As ADODB.Recordset
    Dim sPersCod As String
    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
    If rsProvBS.RecordCount > 0 Then
        Set fgProvBS.Recordset = rsProvBS
    Else
        fgProvBS.Rows = 2
        fgProvBS.Clear
    End If
End Sub
Private Sub Form_Load()
    Dim rsProv  As ADODB.Recordset, rsProvBS  As ADODB.Recordset
    Dim sPersCod As String
    Set clsDProv = New DLogProveedor
    Set clsNProv = New NLogProveedor
    Call CentraForm(Me)
    
    fgProv.Cols = 11
    fgProv.ColWidth(0) = 0
    fgProv.ColWidth(1) = 2400
    fgProv.ColWidth(2) = 1800
    fgProv.ColWidth(3) = 1000
    fgProv.ColWidth(4) = 1000
    'INICIO ORCR-20140913*********
    fgProv.ColWidth(5) = 1000
    fgProv.ColWidth(6) = 0
    fgProv.ColWidth(7) = 1000
    fgProv.ColWidth(8) = 1000
    fgProv.ColWidth(9) = 1000
    fgProv.ColWidth(10) = 1000 'CCI MN PASIERS0472015
    fgProv.ColWidth(11) = 1000 'CCI ME PASIERS0472015
    fgProv.ColWidth(12) = 1000 'Cta Detrac PASIERS0472015
    fgProv.ColWidth(13) = 0
    fgProv.ColWidth(14) = 1500 'Agente RET
    'FIN ORCR-20140913************
    'fgProv.ColWidth(11) = 0
    'fgProv.ColWidth(12) = 1000
    
    fgProvBS.ColWidth(0) = 0
    fgProvBS.ColWidth(1) = 2100
    fgProvBS.ColWidth(2) = 5300
    
    Set rsProv = clsDProv.CargaProveedor
    If rsProv.RecordCount > 0 Then
        Set fgProv.Recordset = rsProv
    Else
        fgProv.TextMatrix(0, 0) = ""
        fgProv.TextMatrix(0, 1) = "Nombre"
        fgProv.TextMatrix(0, 2) = "Dirección"
        fgProv.TextMatrix(0, 3) = "RUC"
        fgProv.TextMatrix(0, 4) = "RUS"
        'INICIO ORCR-20140913*********
        fgProv.TextMatrix(0, 5) = "DOI"
        fgProv.TextMatrix(0, 6) = ""
        fgProv.TextMatrix(0, 7) = "Estado"
        fgProv.TextMatrix(0, 8) = "Cta_MN"
        fgProv.TextMatrix(0, 9) = "Cta_ME"
        fgProv.TextMatrix(0, 10) = "Cta CCI MN" 'PASIERS0472015
        fgProv.TextMatrix(0, 11) = "Cta CCI ME" 'PASIERS0472015
        fgProv.TextMatrix(0, 12) = "Cta Detrac" 'PASIERS0472015
        fgProv.TextMatrix(0, 13) = ""
        fgProv.TextMatrix(0, 14) = "A.Ret/B.Cont"
        'FIN ORCR-20140913************
        'fgProv.TextMatrix(0, 11) = ""
        'fgProv.TextMatrix(0, 12) = "Buen.Cont."
        'cmdProv(0).Enabled = False
        'cmdProv(1).Enabled = False
        'cmdProv(2).Enabled = False
        'cmdProv(3).Enabled = False
        'cmdProv(4).Enabled = False
        'cmdProvBS(0).Enabled = False
        'cmdProvBS(1).Enabled = False
        'Exit Sub
    End If
    
    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
    
    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
    If rsProvBS.RecordCount > 0 Then
        Set fgProvBS.Recordset = rsProvBS
    Else
        fgProvBS.Rows = 2
        fgProvBS.Clear
    End If
    
End Sub

Private Sub ErrorDesc(ByVal pnError As Integer)
    Select Case pnError
        Case 1
            MsgBox "Error al establecer la conexión", vbInformation, " Aviso "
        Case 2
            MsgBox "Registro duplicado", vbInformation, " Aviso "
    End Select
End Sub
