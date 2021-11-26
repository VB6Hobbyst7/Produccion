VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1D5FFCC6-D8BB-11D4-8E95-444553540000}#1.0#0"; "UDSpinText.ocx"
Begin VB.Form frmPaprica 
   Caption         =   "Informacion de Programa PFE"
   ClientHeight    =   6435
   ClientLeft      =   1890
   ClientTop       =   2730
   ClientWidth     =   9960
   Icon            =   "frmpaprica.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9960
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8415
      TabIndex        =   12
      Top             =   5880
      Width           =   1365
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   7995
      TabIndex        =   6
      Top             =   105
      Width           =   1695
   End
   Begin VB.ComboBox cboMeses 
      Height          =   315
      Left            =   885
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   1920
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   697
      TabCaption(0)   =   "Registro de Informe PYMES"
      TabPicture(0)   =   "frmpaprica.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro de Informe de Garantias"
      TabPicture(1)   =   "frmpaprica.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Registro de Informe de Pagos"
      TabPicture(2)   =   "frmpaprica.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Flex"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdGeneraArchaPag"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton CmdGeneraArchaPag 
         Caption         =   "Genera Archivo Pagos"
         Height          =   315
         Left            =   7380
         TabIndex        =   14
         Top             =   4500
         Width           =   1905
      End
      Begin SICMACT.FlexEdit Flex 
         Height          =   3705
         Left            =   180
         TabIndex        =   13
         Top             =   570
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   6535
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cod. IFI-Credito-Cuota-Tipo Pago-Fecha Recuperacion-Saldo Capital-Capital Recuperado-Comision Mensual-Calificacion SBS"
         EncabezadosAnchos=   "1200-2100-1100-1100-1600-1600-1800-2000-2100"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483628
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-2-2"
         TextArray0      =   "Cod. IFI"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   1200
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   9495
         Begin SICMACT.FlexEdit fgGar 
            Height          =   3630
            Left            =   165
            TabIndex        =   11
            Top             =   240
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   6403
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Item-Cod.IFI-Cod.Ope.IFI-TipoGaran-Nom_Garan-Ins.RRPP-Cobertura-Valorizacion-Moneda-Tasacion"
            EncabezadosAnchos=   "450-1200-1800-1200-3500-1200-1200-1200-1200-1200"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-C-L"
            TextArray0      =   "Item"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   450
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdgenArchGar 
            Caption         =   "Generar Arch.Garantias"
            Height          =   375
            Left            =   7440
            TabIndex        =   9
            Top             =   4080
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   5
         Top             =   600
         Width           =   9495
         Begin SICMACT.FlexEdit fgPymes 
            Height          =   3570
            Left            =   120
            TabIndex        =   10
            Top             =   225
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6297
            Cols0           =   33
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   $"frmpaprica.frx":035E
            EncabezadosAnchos=   $"frmpaprica.frx":047A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-15-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   65535
            BackColorControl=   65535
            BackColorControl=   65535
            EncabezadosAlineacion=   "C-L-C-L-L-L-L-C-L-L-L-L-L-C-C-R-R-L-L-L-R-R-C-R-C-L-C-L-L-L-L-C-L"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-2-2-0-0-0-2-3-0-2-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "Item"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   450
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdGenArchPymes 
            Caption         =   "Generar Arch.Pyme"
            Height          =   375
            Left            =   7440
            TabIndex        =   8
            Top             =   3960
            Width           =   1935
         End
      End
   End
   Begin UDSpinText.ucUDSpin txtanio 
      Height          =   345
      Left            =   3285
      TabIndex        =   1
      Top             =   120
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   609
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1998
      MaxLength       =   4
      Text            =   "2002"
      Value           =   2002
   End
   Begin MSComDlg.CommonDialog cmdlOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mes :"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   195
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Año :"
      Height          =   195
      Left            =   2895
      TabIndex        =   3
      Top             =   195
      Width           =   375
   End
End
Attribute VB_Name = "frmPaprica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsApellido As String
Dim gsMaterno As String
Dim gsNombre As String
Private Sub cboMeses_Click()
'txtTC = GetTCFijo(Right(cboMeses, 2), txtanio.value)
End Sub

Private Sub cmdgenArchGar_Click()
 ' Establecer CancelError a True
    If fgGar.TextMatrix(0, 0) = "" Then
        Exit Sub
    End If
    
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    'cmdlOpen.InitDir = App.Path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 3
    
    cmdlOpen.FileName = "GARANTIAS.TXT"
    
    cmdlOpen.ShowSave
    
    Dim lsNombreArchivo As String
    Dim i As Integer
    Dim lsLinea As String
    ' Presentar el nombre del archivo seleccionado
    lsNombreArchivo = cmdlOpen.FileName
    
    Open lsNombreArchivo For Output As #1
    
    For i = 1 To fgPymes.Rows - 1
        lsLinea = fgGar.TextMatrix(i, 1) & _
                    Right(fgGar.TextMatrix(i, 2), 14) & _
                    ImpreFormat(fgGar.TextMatrix(i, 3), 2, 0) & _
                    ImpreFormat(fgGar.TextMatrix(i, 4), 30, 0) & _
                    ImpreFormat("", 10, 0) & _
                    Format(fgGar.TextMatrix(i, 6), "000000000000.00") & _
                    Format(fgGar.TextMatrix(i, 7), "000000000000.00") & _
                    ImpreFormat(fgGar.TextMatrix(i, 8), 2, 0) & _
                    ImpreFormat("", 10, 0)

        Print #1, lsLinea
    Next
    Close #1
    MsgBox "Se ha Generado el archivo " & lsNombreArchivo, vbInformation, "Aviso"
            
    
    
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub cmdGenArchPymes_Click()
    ' Establecer CancelError a True
    If fgPymes.TextMatrix(0, 0) = "" Then
        Exit Sub
    End If
    
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    'cmdlOpen.InitDir = App.Path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 3
    
    cmdlOpen.FileName = "PYMES" & Right(Trim(cboMeses), 2) & txtanio.value & ".TXT"
    
    cmdlOpen.ShowSave
    
    Dim lsNombreArchivo As String
    Dim i As Integer
    Dim lsLinea As String
    ' Presentar el nombre del archivo seleccionado
    lsNombreArchivo = cmdlOpen.FileName
    
    Open lsNombreArchivo For Output As #1
    
    For i = 1 To fgPymes.Rows - 1
        lsLinea = fgPymes.TextMatrix(i, 1) & _
                    Right(fgPymes.TextMatrix(i, 2), 14) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 3), 30, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 4), 30, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 5), 40, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 6), 60, 0) & _
                    fgPymes.TextMatrix(i, 7) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 8), 11, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 9), 60, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 10), 6, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 11), 9, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 12), 1, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 13), 1, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 14), 4, 0) & _
                    Format(fgPymes.TextMatrix(i, 15), "000000000000.00") & _
                    Format(fgPymes.TextMatrix(i, 16), "000000000000.00") & _
                    ImpreFormat(fgPymes.TextMatrix(i, 17), 2, 0) & _
                    Format(fgPymes.TextMatrix(i, 18), "0000") & _
                    ImpreFormat(fgPymes.TextMatrix(i, 19), 2, 0) & _
                    Format(fgPymes.TextMatrix(i, 20), "000000000000.00") & _
                    ImpreFormat(fgPymes.TextMatrix(i, 21), 2, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 22), 1, 0) & _
                    Format(fgPymes.TextMatrix(i, 23), "000000000000.00") & _
                    fgPymes.TextMatrix(i, 24)
            lsLinea = lsLinea & fgPymes.TextMatrix(i, 25) & _
                    fgPymes.TextMatrix(i, 26) & _
                    fgPymes.TextMatrix(i, 27) & _
                    fgPymes.TextMatrix(i, 28) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 29), 60, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 30), 2, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 31), 2, 0) & _
                    ImpreFormat(fgPymes.TextMatrix(i, 32), 1, 0)
                    
        Print #1, lsLinea
    Next
    Close #1
    MsgBox "Se ha Generado el archivo " & lsNombreArchivo, vbInformation, "Aviso"
            
    
    
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub


End Sub

Private Sub CmdGeneraArchaPag_Click()
' Establecer CancelError a True
    If fgGar.TextMatrix(0, 0) = "" Then
        Exit Sub
    End If
    
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    'cmdlOpen.InitDir = App.Path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 3
    
    cmdlOpen.FileName = "PAGOS.TXT"
    
    cmdlOpen.ShowSave
    
    Dim lsNombreArchivo As String
    Dim i As Integer
    Dim lsLinea As String
    ' Presentar el nombre del archivo seleccionado
    lsNombreArchivo = cmdlOpen.FileName
    
    Open lsNombreArchivo For Output As #1
    
    For i = 1 To Flex.Rows - 1
        lsLinea = Flex.TextMatrix(i, 0) & _
                    Right(Flex.TextMatrix(i, 1), 14) & _
                    Flex.TextMatrix(i, 2) & _
                    Flex.TextMatrix(i, 3) & _
                    Flex.TextMatrix(i, 4) & _
                    Format(Flex.TextMatrix(i, 5), "000000000000.00") & _
                    Format(Flex.TextMatrix(i, 6), "000000000000.00") & _
                    Format(Flex.TextMatrix(i, 7), "000000000000.00") & _
                    Flex.TextMatrix(i, 8)

        Print #1, lsLinea
    Next
    Close #1
    MsgBox "Se ha Generado el archivo " & lsNombreArchivo, vbInformation, "Aviso"
            
    
    
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub



Private Sub cmdProcesar_Click()
If cboMeses = "" Then
    MsgBox "Seleccione algun mes del año", vbInformation, "Aviso"
    cboMeses.SetFocus
    Exit Sub
End If

ProcesarPymes Trim(Left(cboMeses, 40)), Trim(Right(cboMeses, 2)), Val(txtanio.value)
ProcesarGarantias Trim(Left(cboMeses, 40)), Trim(Right(cboMeses, 2)), Val(txtanio.value)
ProcesarPagos Trim(Right(cboMeses, 2)), Val(txtanio.value)

End Sub
Sub ProcesarPymes(ByVal lsMes As String, ByVal lnMes As String, ByVal lnAnio As String)
Dim sql As String
Dim rs As ADODB.Recordset
    
sql = "         SELECT  '1000741' AS cCodIFi, PR.cCtaCod as cCodOpeIFI, "
sql = sql & "         CASE WHEN P.nPersPersoneria  = 1 THEN LEFT(P.CPERSNOMBRE,100) ELSE SPACE(100) END AS CNOMCLI, "
sql = sql & "         CASE WHEN P.nPersPersoneria  <> 1 THEN LEFT(P.CPERSNOMBRE,100) ELSE SPACE(100) END AS CRAZON, "
sql = sql & "         cPersIDTpo AS ctidoci, CPERSIDNRO AS cNudoci, P.cPersDireccDomicilio as cdirdom,"
sql = sql & "         left(p.cPersDireccUbiGeo,6) as cCodDom, cPersTelefono as cTelDom,"
sql = sql & "         CASE WHEN P.nPersPersoneria  = 1 THEN 'N' ELSE 'J' END AS CTIPPERS,"
sql = sql & "         cSexo = (select cPersNatSexo from personanat where cPersCod = P.cPersCod),"
sql = sql & "         SUBSTRING(cPersCIIU,2,4) as cCodAct,  000000.00 as nValorAct, 000000.00 as nVentasAnuales,"
sql = sql & "         CASE WHEN nColocDestino = 1 THEN '01' ELSE '02' END as cDesCre, C.nPlazo,"
sql = sql & "         CASE WHEN SUBSTRING(PR.cCtaCod,9,1)='1' THEN '01' ELSE '02' END as cMoneda,"
sql = sql & "         'A' AS cFrec,00000.00 as nComAnual,"
sql = sql & "         dFecSol = (SELECT CONVERT(CHAR(10),dPrdEstado,103) FROM COLOCACESTADO WHERE CCTACOD = PR.CCTACOD AND nPrdEstado = 2000),"
sql = sql & "         dFecApr = (SELECT CONVERT(CHAR(10),dPrdEstado,103) FROM COLOCACESTADO WHERE CCTACOD = PR.CCTACOD AND nPrdEstado = 2002),"
sql = sql & "         nCuoCont= (SELECT nCuotas FROM COLOCACESTADO WHERE CCTACOD = PR.CCTACOD AND nPrdEstado = 2002),"
sql = sql & "         dFecDes as dFecEmiti, dFecDes as dFecVig,"
sql = sql & "         cCodAna = (   SELECT  P1.CPERSNOMBRE"
sql = sql & "                       FROM    PERSONA P1"
sql = sql & "                               JOIN PRODUCTOPERSONA R1 ON R1.CPERSCOD = P1.CPERSCOD"
sql = sql & "                       WHERE  R1.CCTACOD = PR.CCTACOD AND R1.nPrdPersRelac = 28  ),"
sql = sql & "         nCalif = ISNULL((SELECT cCalGen FROM ColocCalifProv where cCtaCod = PR.cCtaCod),''),"
sql = sql & "         '02' AS CTAMEMP,"
sql = sql & "         CASE WHEN SUBSTRING(PR.cCtaCod,6,1)='2' THEN '01' ELSE '02' END AS CTIPMES,"
sql = sql & "         CASE WHEN SUBSTRING(PR.cCtaCod,4,2)='01' THEN '13' ELSE '14' END AS cPFE, DES.nMonto "
sql = sql & "  FROM   PRODUCTO PR"
sql = sql & "         JOIN COLOCACIONES C ON C.CCTACOD = PR.CCTACOD"
sql = sql & "         JOIN COLOCACCRED CC ON CC.CCTACOD = C.CCTACOD"
sql = sql & "         JOIN PRODUCTOPERSONA R ON R.CCTACOD = PR.CCTACOD AND nPrdPersRelac=20"
sql = sql & "         JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD"
sql = sql & "         JOIN (    SELECT  dbo.GetFechaMov(M.CMOVNRO,103) AS dFecDes, MC.cCtaCod, MC.nMonto"
sql = sql & "                   FROM    MOV M"
sql = sql & "                           JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
sql = sql & "                   WHERE   LEFT(M.CMOVNRO,6) =  '" & lnAnio & lnMes & "' AND M.nMovFlag = 0"
sql = sql & "         AND MC.COPECOD LIKE '10010[12345679]%' ) AS DES ON DES.CCTACOD = PR.CCTACOD"
sql = sql & "         LEFT JOIN (SELECT CPERSCOD, CPERSIDNRO,  cPersIDTpo FROM   PERSID) AS CDOC ON CDOC.CPERSCOD = P.CPERSCOD"
sql = sql & "  WHERE   PR.NPRDESTADO IN (2020,2021,2022,2030,2031,2032,2201,2205) AND CC.IDCampana = 15"

Me.fgPymes.Clear
Me.fgPymes.Rows = 2
Me.fgPymes.FormaCabecera
Dim lsTipoDoc As String
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing

rs.Sort = "CNOMCLI"
rs.MoveFirst
Do While Not rs.EOF
    fgPymes.AdicionaFila
    fgPymes.TextMatrix(fgPymes.Row, 1) = rs!cCodIFi
    fgPymes.TextMatrix(fgPymes.Row, 2) = rs!cCodOpeIFI
    gsApellido = ""
    gsMaterno = ""
    gsNombre = ""
    
    DesglosaNombre rs!CNOMCLI
    fgPymes.TextMatrix(fgPymes.Row, 3) = Trim(gsApellido)
    fgPymes.TextMatrix(fgPymes.Row, 4) = Trim(gsMaterno)
    fgPymes.TextMatrix(fgPymes.Row, 5) = Trim(gsNombre)
    fgPymes.TextMatrix(fgPymes.Row, 6) = Trim(rs!cRazon)
    lsTipoDoc = ""
    Select Case rs!ctidoci
        Case "1"
            lsTipoDoc = "01"
        Case "8"
            lsTipoDoc = "01"
        Case "4"
            lsTipoDoc = "02"
        Case "2"
            lsTipoDoc = "03"
        Case "3"
            lsTipoDoc = "04"
        Case "6"
            lsTipoDoc = "05"
        Case "7"
            lsTipoDoc = "06"
    End Select
    fgPymes.TextMatrix(fgPymes.Row, 7) = lsTipoDoc
    fgPymes.TextMatrix(fgPymes.Row, 8) = rs!cnudoci
    fgPymes.TextMatrix(fgPymes.Row, 9) = rs!cDirDom
    fgPymes.TextMatrix(fgPymes.Row, 10) = rs!cCodDom
    fgPymes.TextMatrix(fgPymes.Row, 11) = rs!cTelDom
    fgPymes.TextMatrix(fgPymes.Row, 12) = rs!cTipPers
    fgPymes.TextMatrix(fgPymes.Row, 13) = rs!cSexo
    fgPymes.TextMatrix(fgPymes.Row, 14) = rs!cCodAct
    fgPymes.TextMatrix(fgPymes.Row, 15) = Format(rs!nValorAct, "#0.00")
    fgPymes.TextMatrix(fgPymes.Row, 16) = Format(rs!nVentasAnuales, "#0.00")
    fgPymes.TextMatrix(fgPymes.Row, 17) = rs!cDesCre
    fgPymes.TextMatrix(fgPymes.Row, 18) = rs!nPlazo
    fgPymes.TextMatrix(fgPymes.Row, 19) = rs!cMoneda
    fgPymes.TextMatrix(fgPymes.Row, 20) = Format(rs!nMonto, "#0.00")
    fgPymes.TextMatrix(fgPymes.Row, 21) = rs!nCuoCont
    fgPymes.TextMatrix(fgPymes.Row, 22) = rs!cFrec
    fgPymes.TextMatrix(fgPymes.Row, 23) = Format(rs!nComAnual, "#0.00")
    fgPymes.TextMatrix(fgPymes.Row, 24) = rs!dFecSol
    fgPymes.TextMatrix(fgPymes.Row, 25) = rs!dFecApr
    fgPymes.TextMatrix(fgPymes.Row, 26) = rs!dFecEmiti
    fgPymes.TextMatrix(fgPymes.Row, 27) = rs!dFecVig
    fgPymes.TextMatrix(fgPymes.Row, 28) = DateAdd("d", rs!nPlazo, rs!dFecVig)
    fgPymes.TextMatrix(fgPymes.Row, 29) = rs!cCodAna
    fgPymes.TextMatrix(fgPymes.Row, 30) = rs!cTamEmp
    fgPymes.TextMatrix(fgPymes.Row, 31) = rs!nCalif
    fgPymes.TextMatrix(fgPymes.Row, 32) = rs!cTipMes
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

End Sub

Sub ProcesarGarantias(ByVal lsMes As String, ByVal lnMes As String, ByVal lnAnio As String)
Dim sql As String
Dim rs As ADODB.Recordset

sql = "         SELECT  '1000741' AS cCodIFi, PR.cCtaCod as cCodCta, "
sql = sql & "           G.nTpoGarantia AS CTIPGARAN, G.cDescripcion AS CDESGARA, '' AS DINSRRPP,"
sql = sql & "           G.nRealizacion AS NCOBERTURA, G.nRealizacion AS NVALORIZABIEN,"
sql = sql & "           CASE WHEN G.nMoneda = 1  THEN '01' ELSE '02' END AS CMONEDA,"
sql = sql & "           '' AS DTASACION , G.CNUMGARANT"
sql = sql & "   FROM    PRODUCTO PR"
sql = sql & "           JOIN COLOCACIONES C ON C.CCTACOD = PR.CCTACOD"
sql = sql & "           JOIN COLOCACCRED CC ON CC.CCTACOD = C.CCTACOD"
sql = sql & "           JOIN (  SELECT  dbo.GetFechaMov(M.CMOVNRO,103) AS dFecDes, MC.cCtaCod, MC.nMonto"
sql = sql & "                   FROM    MOV M"
sql = sql & "                           JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
sql = sql & "                   WHERE   LEFT(M.CMOVNRO,6) = '" & lnAnio & lnMes & "' and M.nMovFlag = 0"
sql = sql & "                           AND MC.COPECOD LIKE '10010[12345679]%' ) AS DES ON DES.CCTACOD = PR.CCTACOD"
sql = sql & "           JOIN COLOCGARANTIA CG ON CG.CCTACOD = PR.CCTACOD"
sql = sql & "           JOIN GARANTIAS G on G.cNumGarant = CG.cNumGarant"
sql = sql & "           WHERE   PR.NPRDESTADO IN (2020,2021,2022,2030,2031,2032,2201,2205) AND CC.IDCampana = 15"

Me.fgGar.Clear
Me.fgGar.Rows = 2
Me.fgGar.FormaCabecera
Dim lsTipgar As String
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
rs.Sort = "CCODCTA"
rs.MoveFirst
lsTipgar = ""
Do While Not rs.EOF
    fgGar.AdicionaFila
    fgGar.TextMatrix(fgGar.Row, 1) = rs!cCodIFi
    fgGar.TextMatrix(fgGar.Row, 2) = rs!cCodCta
    lsTipgar = ""
    Select Case Trim(rs!CTIPGARAN)
        Case 1, 19, 25, 30
            lsTipgar = "01"
        Case 2
            lsTipgar = "02"
        Case 4
            lsTipgar = "03"
        Case 22, 24, 26, 27, 28
            lsTipgar = "04"
        Case 5
            lsTipgar = "06"
        Case 23
            lsTipgar = "07"
        Case 2, 29
            lsTipgar = "08"
        Case 8
            lsTipgar = "10"
        Case 6
            lsTipgar = "11"
        Case Else
            lsTipgar = Trim(rs!CTIPGARAN)
    End Select
    fgGar.TextMatrix(fgGar.Row, 3) = lsTipgar
    fgGar.TextMatrix(fgGar.Row, 4) = rs!CDESGARA
    fgGar.TextMatrix(fgGar.Row, 5) = Format(rs!DINSRRPP, "dd/mm/yyyy")
    fgGar.TextMatrix(fgGar.Row, 6) = Format(rs!NCOBERTURA, "#0.00")
    fgGar.TextMatrix(fgGar.Row, 7) = Format(rs!NVALORIZABIEN, "#0.00")
    fgGar.TextMatrix(fgGar.Row, 8) = rs!cMoneda
    fgGar.TextMatrix(fgGar.Row, 9) = Format(rs!DTASACION, "dd/mm/yyyy")
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub


Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
CargaMeses
cboMeses.ListIndex = Month(gdFecSis) - 1
GetTipCambio (gdFecSis)
txtTC = gnTipCambio
txtanio.Text = Year(gdFecSis)
txtMonto = Format(20000, "#,#0.00")
End Sub
Sub CargaMeses()
cboMeses.AddItem "ENERO" & Space(100) & "01"
cboMeses.AddItem "FEBRERO" & Space(100) & "02"
cboMeses.AddItem "MARZO" & Space(100) & "03"
cboMeses.AddItem "ABRIL" & Space(100) & "04"
cboMeses.AddItem "MAYO" & Space(100) & "05"
cboMeses.AddItem "JUNIO" & Space(100) & "06"
cboMeses.AddItem "JULIO" & Space(100) & "07"
cboMeses.AddItem "AGOSTO" & Space(100) & "08"
cboMeses.AddItem "SETIEMBRE" & Space(100) & "09"
cboMeses.AddItem "OCTUBRE" & Space(100) & "10"
cboMeses.AddItem "NOVIEMBRE" & Space(100) & "11"
cboMeses.AddItem "DICIEMBRE" & Space(100) & "12"
End Sub
Public Function DesglosaNombre(psNombre As String) As String
    Dim Total As Long
    Dim Pos As Long
    Dim CadAux As String
    Dim lsApellido As String
    Dim lsMaterno As String
    Dim lsConyugue As String
    Dim lsNombre As String
    
    Total = Len(Trim(psNombre))
    Pos = InStr(psNombre, "/")
    If Pos <> 0 Then
        lsApellido = Left(psNombre, Pos - 1)
        gsApellido = lsApellido
    Else
        CadAux = psNombre
    End If
    
    CadAux = Mid(psNombre, Pos + 1, Total)
    Pos = InStr(CadAux, "\")
    
    If Pos <> 0 Then
        lsMaterno = Left(CadAux, Pos - 1)
        gsMaterno = lsMaterno
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        lsConyugue = Left(CadAux, Pos - 1)
        gsConyugue = lsConyugue
    Else
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos <> 0 Then
            lsMaterno = Left(CadAux, Pos - 1)
            gsMaterno = lsMaterno
            lsConyugue = ""
            gsConyugue = lsConyugue
        End If
    End If
    lsNombre = Mid(CadAux, Pos + 1, Total)
    gsNombre = lsNombre
    DesglosaNombre = lsApellido & " " & IIf(Len(Trim(lsConyugue)) = 0, lsMaterno, lsConyugue) & " " & lsNombre
End Function

Sub ProcesarPagos(ByVal psMes As String, ByVal psAno As String)
    Dim sSQL As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim nFila As Integer
    
    sSQL = "Select '1000741' AS cCodIFi,MD.cCtaCod,MD.nNroCuota,'01' as CTipoPago,"
    sSQL = sSQL & "SubString(M.cMovNro,7,2)+'/'+SubString(M.cMovNro,5,2)+'/'+SubString(M.cMovNro,1,4) as dFecha,"
    sSQL = sSQL & " P.nSaldo as SK,"
    sSQL = sSQL & "MD.nMonto as nMontoRecup,"
    sSQL = sSQL & " 0 as nMontoComision,"
    sSQL = sSQL & " Cp.cCalGen"
    sSQL = sSQL & "  From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro and MD.nPrdConceptoCod=1000"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=MD.cCtaCod"
    sSQL = sSQL & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSQL = sSQL & " Left Join ColocCalifProv CP on CP.cPerscod=PP.cPersCod and CP.cCtaCod=PP.cCtaCod"
    sSQL = sSQL & " Inner Join ColocacCred C on C.cCtaCod=P.cCtaCod"
    sSQL = sSQL & " Where Left(M.cMovNro,6)='" & psAno & psMes & "'  and M.nMovFlag=0 and C.IDCampana=15 and MD.cOpeCod not Like '107[123456789]%'"
    sSQL = sSQL & " Order By MD.cCtaCod"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Flex.Clear
    Flex.FormaCabecera
    
    nFila = 1
    Do Until rs.EOF
        Flex.AdicionaFila
        Flex.TextMatrix(nFila, 0) = rs!cCodIFi
        Flex.TextMatrix(nFila, 1) = rs!cCtaCod
        Flex.TextMatrix(nFila, 2) = IIf(rs!nNroCuota > 9, rs!nNroCuota, "0" & rs!nNroCuota)
        Flex.TextMatrix(nFila, 3) = rs!CTipoPago
        Flex.TextMatrix(nFila, 4) = rs!dFecha
        Flex.TextMatrix(nFila, 5) = Format(rs!SK, "#0.00")
        Flex.TextMatrix(nFila, 6) = Format(rs!nMontoRecup, "#0.00")
        Flex.TextMatrix(nFila, 7) = Format(rs!nMontoComision, "#0.00")
        Flex.TextMatrix(nFila, 8) = IIf(IsNull(rs!cCalGen), 0, rs!cCalGen)
        nFila = nFila + 1
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
