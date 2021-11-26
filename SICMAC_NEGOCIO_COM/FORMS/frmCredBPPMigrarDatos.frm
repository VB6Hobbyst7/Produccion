VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPMigrarDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Migrar Datos de Mes Anterior"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmCredBPPMigrarDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Migrar Configuración Mensual"
      TabPicture(0)   =   "frmCredBPPMigrarDatos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMesAnio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "feParametrosGen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbAgencias"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdMigrar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCerrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdMigrar 
         Caption         =   "Migrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   3960
         Width           =   1335
      End
      Begin VB.ComboBox cmbAgencias 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin SICMACT.FlexEdit feParametrosGen 
         Height          =   2610
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   6060
         _extentx        =   10689
         _extenty        =   4604
         cols0           =   5
         highlight       =   2
         allowuserresizing=   3
         encabezadosnombres=   "N°-Parametro-Aplicar-Mes/Año-CodParam"
         encabezadosanchos=   "0-3000-800-1800-0"
         font            =   "frmCredBPPMigrarDatos.frx":0326
         font            =   "frmCredBPPMigrarDatos.frx":034E
         font            =   "frmCredBPPMigrarDatos.frx":0376
         font            =   "frmCredBPPMigrarDatos.frx":039E
         font            =   "frmCredBPPMigrarDatos.frx":03C6
         fontfixed       =   "frmCredBPPMigrarDatos.frx":03EE
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-4-3-0"
         encabezadosalineacion=   "C-L-L-L-C"
         formatosedit    =   "0-0-0-0-0"
         avanceceldas    =   1
         textarray0      =   "N°"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
      End
      Begin VB.Label lblMesAnio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes-Año:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   800
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCredBPPMigrarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'Private fgFecActual As Date
'
'Private Sub cmbAgencias_Click()
'CargaParametros
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub CargaControles()
'MesActual
'CargaComboAgencias cmbAgencias
'CargaParametros
'End Sub
'Private Sub MesActual()
'Dim oConsSist As COMDConstSistema.NCOMConstSistema
'Set oConsSist = New COMDConstSistema.NCOMConstSistema
'fgFecActual = oConsSist.LeeConstSistema(gConstSistFechaBPP)
'Set oConsSist = Nothing
'
'lblMesAnio.Caption = MesAnio(fgFecActual)
'End Sub
'Private Function MesAnio(ByVal dFecha As Date) As String
'Dim sFechaDesc As String
'sFechaDesc = ""
'
'Select Case Month(dFecha)
'    Case 1: sFechaDesc = "Enero"
'    Case 2: sFechaDesc = "Febrero"
'    Case 3: sFechaDesc = "Marzo"
'    Case 4: sFechaDesc = "Abril"
'    Case 5: sFechaDesc = "Mayo"
'    Case 6: sFechaDesc = "Junio"
'    Case 7: sFechaDesc = "Julio"
'    Case 8: sFechaDesc = "Agosto"
'    Case 9: sFechaDesc = "Septiembre"
'    Case 10: sFechaDesc = "Octubre"
'    Case 11: sFechaDesc = "Noviembre"
'    Case 12: sFechaDesc = "Diciembre"
'End Select
'
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
'Private Sub CargaParametros()
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'Dim i As Integer
'
'Set oConst = New COMDConstantes.DCOMConstantes
'Set rsConst = oConst.RecuperaConstantes(7065)
'
'LimpiaFlex feParametrosGen
'
'If Not (rsConst.EOF And rsConst.BOF) Then
'    For i = 0 To rsConst.RecordCount - 1
'        feParametrosGen.AdicionaFila
'        feParametrosGen.TextMatrix(i + 1, 2) = ""
'        feParametrosGen.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'        feParametrosGen.TextMatrix(i + 1, 4) = Trim(rsConst!nConsValor)
'        rsConst.MoveNext
'    Next i
'End If
'feParametrosGen.TopRow = 1
'
'Set oConst = Nothing
'Set rsConst = Nothing
'End Sub
'
'Private Sub cmdMigrar_Click()
'If ValidaDatos Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Dim lnAnioMigrar As Integer
'    Dim lnMesMigrar As Integer
'    Dim lsAux As String
'    Dim lsCodAge As String
'    Dim lnTipo As Integer
'    lsCodAge = Trim(Right(cmbAgencias.Text, 5))
'
'    For i = 1 To feParametrosGen.Rows - 1
'        If Trim(feParametrosGen.TextMatrix(i, 2)) = "." Then
'            lnTipo = CInt(feParametrosGen.TextMatrix(i, 4))
'
'            lsAux = Trim(Right(feParametrosGen.TextMatrix(i, 3), 6))
'            lnAnioMigrar = CInt(Right(lsAux, 4))
'
'            If Len(lsAux) = 6 Then
'                lnMesMigrar = CInt(Left(lsAux, 2))
'            Else
'                lnMesMigrar = CInt(Left(lsAux, 1))
'            End If
'
'            Call oBPP.MigrarDatosXParam(lnMesMigrar, lnAnioMigrar, lsCodAge, lnTipo, Month(fgFecActual), Year(fgFecActual))
'
'        End If
'    Next
'
'    MsgBox "Se migraron los datos Satisfactoriamente", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Function ValidaDatos() As Boolean
'Dim nCont As Integer
'If Trim(cmbAgencias.Text) = "" Then
'    MsgBox "Selecciones la Agencia", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'nCont = 0
'For i = 1 To feParametrosGen.Rows - 1
'    If Trim(feParametrosGen.TextMatrix(i, 2)) = "." Then
'        nCont = nCont + 1
'        If Trim(feParametrosGen.TextMatrix(i, 3)) = "" Then
'            MsgBox "Selecciones el Mes que se va a Migrar", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'    End If
'Next
'
'If nCont = 0 Then
'    MsgBox "Selecciones por lo menos un Parametro a Migrar", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'ValidaDatos = True
'End Function
'
'Private Sub feParametrosGen_OnCellChange(pnRow As Long, pnCol As Long)
'    Dim i As Integer
'    i = 0
'    For i = 1 To Me.feParametrosGen.Rows - 1
'        If feParametrosGen.TextMatrix(i, 2) <> "." Then
'            feParametrosGen.TextMatrix(i, 3) = ""
'        End If
'    Next
'End Sub
'
'Private Sub feParametrosGen_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'If feParametrosGen.lbEditarFlex Then
'    Select Case feParametrosGen.Col
'        Case 2
'            If Trim(feParametrosGen.TextMatrix(feParametrosGen.row, 2)) = "." Then
'                feParametrosGen.ColumnasAEditar = "X-X-2-3-X-X"
'            Else
'                feParametrosGen.ColumnasAEditar = "X-X-2-X-X-X"
'                feParametrosGen.TextMatrix(feParametrosGen.row, 3) = ""
'            End If
'    End Select
'End If
'End Sub
'
'Private Sub feParametrosGen_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'Dim sColumnas() As String
'sColumnas = Split(feParametrosGen.ColumnasAEditar, "-")
'If sColumnas(pnCol) = "X" Then
'   Cancel = False
'   SendKeys "{Tab}", True
'   Exit Sub
'End If
'End Sub
'
'Private Sub feParametrosGen_RowColChange()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim lsCodAge As String
'Set oBPP = New COMNCredito.NCOMBPPR
'
'lsCodAge = ""
'If Trim(cmbAgencias.Text) = "" Then
'    MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'    Exit Sub
'End If
'
'lsCodAge = Trim(Right(cmbAgencias.Text, 4))
'If feParametrosGen.lbEditarFlex Then
'    Select Case feParametrosGen.Col
'        Case 3
'            If Trim(feParametrosGen.TextMatrix(feParametrosGen.row, 2)) = "." Then
'                feParametrosGen.ColumnasAEditar = "X-X-2-3-X-X"
'                feParametrosGen.CargaCombo oBPP.DevolverFechasAMigrarXParam(CInt(feParametrosGen.TextMatrix(feParametrosGen.row, 4)), Month(fgFecActual), Year(fgFecActual), lsCodAge)
'            Else
'                feParametrosGen.ColumnasAEditar = "X-X-2-X-X-X"
'            End If
'    End Select
'End If
'End Sub
'
'Private Sub Form_Load()
'CargaControles
'End Sub
'
'
