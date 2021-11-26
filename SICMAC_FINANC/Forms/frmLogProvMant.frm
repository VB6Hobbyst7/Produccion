VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProvMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   5925
   ClientLeft      =   960
   ClientTop       =   1170
   ClientWidth     =   9240
   Icon            =   "frmLogProvMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdProv 
      Caption         =   "Imprimir"
      Height          =   360
      Index           =   4
      Left            =   7305
      TabIndex        =   11
      Top             =   2775
      Width           =   1500
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Buscar"
      Height          =   360
      Index           =   3
      Left            =   5685
      TabIndex        =   4
      Top             =   2775
      Width           =   1500
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Eliminar"
      Height          =   360
      Index           =   2
      Left            =   4035
      TabIndex        =   3
      Top             =   2775
      Width           =   1500
   End
   Begin VB.CommandButton cmdProvBS 
      Caption         =   "Eliminar B/S"
      Height          =   360
      Index           =   1
      Left            =   2460
      TabIndex        =   7
      Top             =   5445
      Width           =   1500
   End
   Begin VB.CommandButton cmdProvBS 
      Caption         =   "Agregar B/S"
      Height          =   360
      Index           =   0
      Left            =   750
      TabIndex        =   6
      Top             =   5445
      Width           =   1500
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Activar/Desactivar"
      Height          =   360
      Index           =   1
      Left            =   2220
      TabIndex        =   2
      Top             =   2775
      Width           =   1500
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Agregar"
      Height          =   360
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   2775
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7200
      TabIndex        =   8
      Top             =   5415
      Width           =   1305
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProv 
      Height          =   2310
      Left            =   225
      TabIndex        =   0
      Top             =   375
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4075
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
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
      Height          =   1680
      Left            =   195
      TabIndex        =   5
      Top             =   3600
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   2963
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   270
      X2              =   8970
      Y1              =   3225
      Y2              =   3225
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   255
      X2              =   8955
      Y1              =   3210
      Y2              =   3210
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
      Left            =   390
      TabIndex        =   10
      Top             =   3315
      Width           =   2640
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Proveedores"
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
      Index           =   0
      Left            =   390
      TabIndex        =   9
      Top             =   135
      Width           =   1365
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

Private Sub cmdProv_Click(Index As Integer)
    Dim rsProvBS As ADODB.Recordset
    Dim obp As UPersona
    Dim nCont As Integer, nResult As Integer
    Dim sActualiza As String
    Dim sPersCod As String
    Dim sPersNombre As String
    Dim nProvEstado As Integer
    Dim bPersDoc As Boolean
    
    If Index > 0 Then
        If fgProv.Row = 0 Then
            MsgBox "No existen proveedores", vbInformation, " Aviso"
            Exit Sub
        End If
        sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
        sPersNombre = fgProv.TextMatrix(fgProv.Row, 1)
        nProvEstado = Val(fgProv.TextMatrix(fgProv.Row, 5))
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
                bPersDoc = False
                For nCont = 1 To obp.DocsPers.RecordCount
                    If obp.DocsPers!cPersIdTpo = gPersIdRUS Then
                        bPersDoc = True
                        Exit For
                    End If
                    obp.DocsPers.MoveNext
                Next
            End If
            If sPersCod <> "" And bPersDoc Then
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nResult = clsDProv.GrabaProveedor(sPersCod, sActualiza)
                If nResult = 0 Then
                    Set fgProv.Recordset = clsDProv.CargaProveedor
                    sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
                    
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
                If nResult = 0 Then
                    Set fgProv.Recordset = clsDProv.CargaProveedor
                    sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
                    
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
        Case 3:
            'BUSCAR
            sPersNombre = Trim(frmLogMantOpc.Inicio(0, 1))
            If sPersNombre <> "" Then
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
        Case Else
            MsgBox "Indice de comand de proveedores no reconocido", vbInformation, " Aviso "
    End Select
    
    frmMdiMain.staMain.Panels(2).Text = fgProv.Rows & " registros."
End Sub

Private Sub cmdProvBS_Click(Index As Integer)
    Dim clsDBS As DLogBieSer
    Dim rsProvBS As ADODB.Recordset
    Dim nResult As Integer
    Dim sActualiza As String
    Dim sPersCod As String
    Dim sBSCod As String
    Dim sBSNombre As String
    
    If fgProv.Row = 0 Then
        MsgBox "No existen proveedores", vbInformation, " Aviso"
        Exit Sub
    End If
    sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
    
    If Index > 0 Then
        If fgProvBS.Row = 0 Then
            MsgBox "No existen bienes/servicios del proveedor", vbInformation, " Aviso"
            Exit Sub
        End If
        sBSCod = fgProvBS.TextMatrix(fgProvBS.Row, 0)
        sBSNombre = fgProvBS.TextMatrix(fgProvBS.Row, 2)
    End If
    
    Select Case Index
        Case Is = 0
            'Agregar bien/servicio de proveedor
            Dim vBS As ClassDescObjeto
            Set vBS = New ClassDescObjeto
            Set clsDBS = New DLogBieSer
            
            vBS.Show clsDBS.CargaBS(1), ""
            
            Set clsDBS = Nothing
            If vBS.lbOk Then
                sBSCod = vBS.gsSelecCod
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                
                nResult = clsDProv.GrabaProveedorBS(sPersCod, sBSCod, sActualiza)
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
    sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
    If rsProvBS.RecordCount > 0 Then
        Set fgProvBS.Recordset = rsProvBS
    Else
        fgProvBS.Rows = 2
        fgProvBS.Clear
    End If
End Sub

Private Sub Form_Load()
    Dim rsProvBS  As ADODB.Recordset
    Dim sPersCod As String
    Set clsDProv = New DLogProveedor
    Set clsNProv = New NLogProveedor
    Call CentraForm(Me)
    
    fgProv.ColWidth(0) = 0
    fgProv.ColWidth(1) = 2500
    fgProv.ColWidth(2) = 2500
    fgProv.ColWidth(3) = 1100
    fgProv.ColWidth(4) = 1100
    fgProv.ColWidth(5) = 0
    fgProv.ColWidth(6) = 1000
    
    fgProvBS.ColWidth(0) = 0
    fgProvBS.ColWidth(1) = 2000
    fgProvBS.ColWidth(2) = 6000
    
    Set fgProv.Recordset = clsDProv.CargaProveedor
    sPersCod = fgProv.TextMatrix(fgProv.Row, 0)
    
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
