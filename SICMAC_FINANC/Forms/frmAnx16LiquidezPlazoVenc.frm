VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnx16LiquidezPlazoVenc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexo 16: Liquidez por Plazos de Vecimiento"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   Icon            =   "frmAnx16LiquidezPlazoVenc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   4800
      Width           =   1395
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   4800
      Width           =   1395
   End
   Begin VB.Frame fraPC 
      Caption         =   "Plan de Contingencia"
      ForeColor       =   &H8000000D&
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   8895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   7680
         TabIndex        =   24
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   330
         Left            =   6600
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin Sicmact.FlexEdit grdContingencia 
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Entidad-Total MN-Total ME"
         EncabezadosAnchos=   "500-4200-1500-1500"
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
         ColumnasAEditar =   "X-1-2-3"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R"
         FormatosEdit    =   "0-0-2-2"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   5160
      TabIndex        =   15
      Top             =   120
      Width           =   3855
      Begin Sicmact.EditMoney txtMora 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin Sicmact.EditMoney txtPatEfec 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Pat.Efe:"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "% Mora:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "TC:"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Año:"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraRep 
      Caption         =   "Encaje"
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      Begin Sicmact.EditMoney txtEncajeMN 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin Sicmact.EditMoney txtEncajeME 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólares:"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Soles:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAnx16LiquidezPlazoVenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsFecha As String
Dim dPatrimonio As Date
Dim sOpeReporte As String


Private Sub cmbMes_Change()
    txtTipCambio.Text = TipoCambioCierre(mskAnio, cmbMes.ListIndex + 1)
End Sub

Private Sub cmbMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskAnio.SetFocus
    End If
End Sub

Private Sub cmdAgregar_Click()
    Dim nFila As Long
    Dim Entidades() As String
   
    Me.grdContingencia.AdicionaFila
    nFila = Me.grdContingencia.Rows - 1
    Me.grdContingencia.SetFocus
    SendKeys "{Enter}"
    Me.grdContingencia.TextMatrix(nFila, 1) = ""
    Me.grdContingencia.TextMatrix(nFila, 2) = "0.00"
    Me.grdContingencia.TextMatrix(nFila, 3) = "0.00"
End Sub
Private Sub cmdGenerar_Click()
    If ValidaDatos Then
           generarAnexo
    End If
End Sub
Private Function ValidaDatos() As Boolean
    Dim Conecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim ldFecha As Date
    Dim dfecha As Date
    Dim sql As String
    
    Set Conecta = New DConecta
    If mskAnio.Text = "" Then
        MsgBox "Debe ingresar el año", vbInformation, "SICMACM"
        mskAnio.SetFocus
        Exit Function
    End If
    Conecta.AbreConexion
    sql = "Select dfecConsol From DBConsolidada..VarConsolida"
    Set rs = Conecta.CargaRecordSet(sql)
    dfecha = rs!dfecConsol
    Set rs = Nothing
    ldFecha = CDate(mskAnio.Text & "/" & Format(Right(cmbMes.Text, 2), "00") & "/01")
    ldFecha = DateAdd("m", 1, ldFecha)
    ldFecha = DateAdd("d", -1, ldFecha)
    If DateDiff("m", ldFecha, dfecha) <> 0 Then
        MsgBox "El periodo ingresado debe corresponder al último cierre mensual", vbInformation, "SICMACM"
        mskAnio.SetFocus
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Sub generarAnexo()
    Dim sPathAnexo16 As String
    Dim fs As New Scripting.FileSystemObject
    Dim obj_Excel As Object, Libro As Object, Hoja As Object
    Dim ldFecha As Date
    Dim i, j, m, N As Integer
    
    On Error GoTo error_sub
    
    ldFecha = CDate(mskAnio.Text & "/" & Format(Right(cmbMes.Text, 2), "00") & "/01")
    ldFecha = DateAdd("m", 1, ldFecha)
    ldFecha = DateAdd("d", -1, ldFecha)
          
    If sOpeReporte = "770162" Then
        sPathAnexo16 = App.path & "\Spooler\Anexo16B_" + Format(ldFecha, "yyyymmdd") + ".xls"
    Else
        sPathAnexo16 = App.path & "\Spooler\Anexo16_" + Format(ldFecha, "yyyymmdd") + ".xls"
    End If
    
    If fs.FileExists(sPathAnexo16) Then
        If ArchivoEstaAbierto(sPathAnexo16) Then
            If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo16) + " para continuar", vbRetryCancel) = vbCancel Then
                Me.MousePointer = vbDefault
                Exit Sub
            End If
                Me.MousePointer = vbHourglass
            End If
            fs.DeleteFile sPathAnexo16, True
        End If
              
        If sOpeReporte = "770162" Then
            sPathAnexo16 = App.path & "\FormatoCarta\Anexo16B_SBS.xls"
        Else
            sPathAnexo16 = App.path & "\FormatoCarta\Anexo16_SBS.xls"
        End If
        
        If Len(Dir(sPathAnexo16)) = 0 Then
            MsgBox "No se Pudo Encontrar el Archivo:" & sPathAnexo16, vbCritical
            Exit Sub
        End If
        
        Set obj_Excel = CreateObject("Excel.Application")
        obj_Excel.DisplayAlerts = False
        Set Libro = obj_Excel.Workbooks.Open(sPathAnexo16)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Dim oCtaCont As dBalanceCont
        Dim rsCtaCont As ADODB.Recordset
        
        Set oCtaCont = New dBalanceCont
        Set rsCtaCont = New ADODB.Recordset
       ' Fecha del ANEXO
        Set celda = IIf(sOpeReporte = "770160", obj_Excel.Range("A4"), obj_Excel.Range("A5"))
        celda.value = Format(ldFecha, "Al dd MMMM yyyy")
        
        If sOpeReporte = "770162" Then
            sPathAnexo16 = App.path & "\Spooler\ANEXO_16B_" + Format(ldFecha, "yyyymmdd") + ".xls"
        Else
            sPathAnexo16 = App.path & "\Spooler\ANEXO_16_" + Format(ldFecha, "yyyymmdd") + ".xls"
        End If
        If fs.FileExists(sPathAnexo16) Then
            If ArchivoEstaAbierto(sPathAnexo16) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo16)
            End If
            fs.DeleteFile sPathAnexo16, True
        End If
       
        '********************************** ACTIVOS *******************************************************
        ' 1100 - Activo Disponible
        Call ActivoDisponible(ldFecha, obj_Excel)
        ' 1300 - Inversiones Negociables y a Vecimiento
        Set celda = IIf(sOpeReporte = "770162", obj_Excel.Range("L12"), obj_Excel.Range("N12"))
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("13", ldFecha, "2", Me.txtTipCambio.Text) '"30420.89"
        ' 1401,1403 - Créditos
        Call ActivoCreditos(ldFecha, obj_Excel)
        ' 1500 - Cuentas por cobrar
        Set celda = IIf(sOpeReporte = "770162", obj_Excel.Range("K14"), obj_Excel.Range("M14"))
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("15", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = IIf(sOpeReporte = "770162", obj_Excel.Range("L14"), obj_Excel.Range("N14"))
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("15", ldFecha, "2", Me.txtTipCambio.Text)
        ' 1601,1602 - Bienes realizables
        Set celda = IIf(sOpeReporte = "770162", obj_Excel.Range("K15"), obj_Excel.Range("M15"))
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1612", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = IIf(sOpeReporte = "770162", obj_Excel.Range("L15"), obj_Excel.Range("N15"))
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1622", ldFecha, "2", Me.txtTipCambio.Text)
        
        '********************************* PASIVOS *******************************************************
        ' 2101 - Obligaciones a la Vista
        Set celda = obj_Excel.Range("C20")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2111", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = obj_Excel.Range("D20")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2121", ldFecha, "1", Me.txtTipCambio.Text)
        '2102 - Obligaciones por Cuentas de Ahorro
        Call PasivoDepositoAhorros(ldFecha, obj_Excel)
        '2102 - Obligaciones por Cuentas a Plazo
        Call PasivoDepositoPlazo(ldFecha, obj_Excel)
        
        ' 2104,2106,2107,2108 - Obligaciones con el publico
        Set celda = obj_Excel.Range("C24")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2114,2116,2118", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = obj_Excel.Range("D24")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2124,2126,2128", ldFecha, "2", Me.txtTipCambio.Text)
        Set celda = obj_Excel.Range("K24")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2117", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = obj_Excel.Range("L24")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2127", ldFecha, "2", Me.txtTipCambio.Text)
        
        '2300 - Depositos de Ifis y Ofi
        Call DepositosIFISyOFI(ldFecha, obj_Excel)
        'Adeudados y obligaciones financieras
        Call cargarCalecAdeudados(ldFecha, obj_Excel)
        'Cuentas por pagar
        Set celda = obj_Excel.Range("C28")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("25", ldFecha, "1", Me.txtTipCambio.Text)
        Set celda = obj_Excel.Range("D28")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("25", ldFecha, "2", Me.txtTipCambio.Text)
        
        '**************************** BRECHA ACUMULADA(III) / PATRIMONIO EFECTIVO *********************************
        If sOpeReporte = "770162" Then
            i = 36
        Else
            i = 40
        End If
        
        Set celda = obj_Excel.Range("C" & i)
        celda.value = "=C35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("D" & i)
        celda.value = "=D35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("E" & i)
        celda.value = "=E35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("F" & i)
        celda.value = "=F35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("G" & i)
        celda.value = "=G35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("H" & i)
        celda.value = "=H35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("I" & i)
        celda.value = "=I35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("J" & i)
        celda.value = "=J35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("K" & i)
        celda.value = "=K35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("L" & i)
        celda.value = "=L35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("M" & i)
        celda.value = "=M35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("N" & i)
        celda.value = "=N35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("O" & i)
        celda.value = "=O35/" & CStr(CDbl(Me.txtPatEfec.Text))
        Set celda = obj_Excel.Range("P" & i)
        celda.value = "=P35*" & Me.txtTipCambio.Text & "/" & CStr(CDbl(Me.txtPatEfec.Text))
        
        '******************************** INDICADORES (Anexo 16) ************************************
        If sOpeReporte = "770160" Then
            Call ObtenerIndicadores(obj_Excel, ldFecha)
        End If
        '*************************** PLAN DE CONTINGENCIA (Anexo 16B) ****************************
        If sOpeReporte = "770162" Then
            Call PlanDeContingencia(obj_Excel)
        End If
        
        Hoja.SaveAs sPathAnexo16
        'obj_Excel.Visible = True
        Libro.Close
        obj_Excel.Quit
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_Excel = Nothing
        Me.MousePointer = vbDefault
        'abre y muestra el archivo
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathAnexo16)
        m_excel.Visible = True

Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_Excel = Nothing
        Set Hoja = Nothing
        'PB1.Visible = False
        'lblAvance.Caption = ""
        Me.MousePointer = vbDefault
    
End Sub
Private Sub ObtenerIndicadores(ByVal obj_Excel As Excel.Application, ByVal ldFecha As Date)
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim oCtaCont As dBalanceCont
    Dim nTop20CapPrinCli As Double
    Dim nTop10CapPrinCli As Double
    Dim nTop05CapPrinCli As Double
    Dim nTop10Acreedores As Double
    Dim nTotalDepositos As Double
    Dim nTotalDepEntidadesPub As Double
    Dim nPasivoTotal As Double
    Dim celda As Excel.Range
    Dim i As Integer
    
    Set oCtaIf = New NCajaCtaIF
    Set oCtaCont = New dBalanceCont
            
    Set celda = obj_Excel.Range("K43")
    celda.value = "=Round((M13 + N13 * " & CStr(CDbl(Me.txtTipCambio.Text)) & ")/" & CStr(CDbl(Me.txtPatEfec.Text)) & ",2)"
    Set celda = obj_Excel.Range("K44")
    celda.value = "=Round(((M21+M22+M26) + (N21+N22+N26) * " & CStr(CDbl(Me.txtTipCambio.Text)) & ")/" & CStr(CDbl(Me.txtPatEfec.Text)) & ",2)"
    Set celda = obj_Excel.Range("K45")
    celda.value = "=K43 - K44"
    
    Set prs = oCtaIf.GetCapPrincipalesClientes(20, CDbl(Me.txtTipCambio.Text))
    While Not prs.EOF And Not prs.BOF
        i = i + 1
        If i <= 5 Then
            nTop05CapPrinCli = nTop05CapPrinCli + prs!nSaldo
        End If
        If i <= 10 Then
            nTop10CapPrinCli = nTop10CapPrinCli + prs!nSaldo
        End If
        nTop20CapPrinCli = nTop20CapPrinCli + prs!nSaldo
        prs.MoveNext
    Wend
    Set prs = Nothing
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("2412", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2422", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("2612", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2622", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("2614020101", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2624020101", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("26170101020101", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("26270101020101", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("24130201010105", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("24230201010105", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("24130201010131", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("24230201010131", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("24130201010129", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("24230201010129", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("241601010102", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("240601010102", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("2616020101", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2626020101", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("261501010101", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("262501010101", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTop10Acreedores = nTop10Acreedores + oCtaCont.ObtenerCtaContBalanceMensual("2415010101", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2425010101", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    nTop10Acreedores = nTop10Acreedores + nTop05CapPrinCli
    
    nTotalDepositos = oCtaIf.GetCapTotalDepositos(Format(ldFecha, "yyyyMMdd"), CDbl(Me.txtTipCambio.Text))
    
    nPasivoTotal = oCtaCont.ObtenerCtaContBalanceMensual("2", ldFecha, "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("2", ldFecha, "2", Me.txtTipCambio.Text) * CDbl(Me.txtTipCambio.Text)
    
    nTotalDepEntidadesPub = oCtaIf.GetCapDepositosEntidadesPublicas(CDbl(Me.txtTipCambio.Text))
    
    Set celda = obj_Excel.Range("D43")
    celda.value = Format(Round(nTop10Acreedores / nPasivoTotal, 2), "0.00")
    Set celda = obj_Excel.Range("D44")
    celda.value = Format(Round(nTop10CapPrinCli / nTotalDepositos, 2), "0.00")
    Set celda = obj_Excel.Range("D45")
    celda.value = Format(Round(nTop20CapPrinCli / nTotalDepositos, 2), "0.00")
    Set celda = obj_Excel.Range("D46")
    celda.value = Format(Round(nTotalDepEntidadesPub / nTotalDepositos, 2), "0.00")
    
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
End Sub
Private Sub PlanDeContingencia(ByVal obj_Excel As Excel.Application)
    Dim nFila, nFilaIni As Integer
    Dim i, j As Integer
    Dim nMontosLineas(1 To 20, 1 To 2) As Double
    Dim nPivotMN As Double: Dim nPivotME As Double
    Dim nTotAcumMN As Double: Dim nTotAcumME As Double
    Dim celda As Excel.Range
    
    nFila = 39
    nFilaIni = nFila
    For i = 1 To Me.grdContingencia.Rows - 1
        nMontosLineas(i, 1) = Me.grdContingencia.TextMatrix(i, 2)
        nMontosLineas(i, 2) = Me.grdContingencia.TextMatrix(i, 3)
       
        Set celda = obj_Excel.Range("B" & nFila)
        celda.value = Me.grdContingencia.TextMatrix(i, 1)
        Set celda = obj_Excel.Range("M" & nFila)
        celda.value = Replace("=C[X]+E[X]+G[X]+I[X]+K[X]", "[X]", CStr(nFila))
        Set celda = obj_Excel.Range("N" & nFila)
        celda.value = Replace("=D[X]+F[X]+H[X]+J[X]+L[X]", "[X]", CStr(nFila))
        nFila = nFila + 1
    Next
    
    Call EstableceFormulas(obj_Excel, nFila, "PARA GESTIONAR - FINANZAS", "")
    Call EstableceFormulas(obj_Excel, nFila + 1, "TOTAL (IV)", "=SUM([X]39:[X]" & nFila & ")")
    Call EstableceFormulas(obj_Excel, nFila + 3, "TOTAL (I)-(II)+(IV)", "=[X]34+[X]" & nFila + 1)
    Call EstableceFormulas(obj_Excel, nFila + 4, "TOTAL ACUMULADO", "=IF([X]8=""M.N."",SUMIF(C8:[X]8,""M.N."",C" & nFila + 3 & ":[X]" & nFila + 3 & "),SUMIF(C8:[X]8,""M.E."",C" & nFila + 3 & ":[X]" & nFila + 3 & "))")
    Call EstableceFormulas(obj_Excel, nFila + 5, "TOTAL ACUMULADO/PATRIM.EFECTIVO", "=[X]" & nFila + 3 & "/" & CStr(CDbl(Me.txtPatEfec.Text)))

    For i = 1 To Me.grdContingencia.Rows - 1
        nPivotMN = Me.grdContingencia.TextMatrix(i, 2) 'Monto total de la linea en MN
        nPivotME = Me.grdContingencia.TextMatrix(i, 3) 'Monto total de la linea en ME
        For j = 1 To 5
            nTotAcumMN = obj_Excel.Cells(nFila + 4, j * 2 + 1)
            nTotAcumME = obj_Excel.Cells(nFila + 4, j * 2 + 2)

            If nTotAcumMN < 0 And nPivotMN > 0 Then
                Set celda = obj_Excel.Cells(nFilaIni + i - 1, j * 2 + 1)
                If nTotAcumMN + nPivotMN < 0 Then
                    celda.value = nPivotMN
                    nPivotMN = 0
                Else
                    celda.value = Abs(nTotAcumMN)
                    nPivotMN = nPivotMN + nTotAcumMN
                End If
                
            End If
            If nTotAcumME < 0 And nPivotME > 0 Then
                Set celda = obj_Excel.Cells(nFilaIni + i - 1, j * 2 + 2)
                If nTotAcumME + nPivotME < 0 Then
                    celda.value = nPivotME
                    nPivotME = 0
                Else
                    celda.value = Abs(nTotAcumME)
                    nPivotME = nPivotME + nTotAcumME
                End If
            End If
        Next
    Next
    For j = 1 To 5
        nTotAcumMN = obj_Excel.Cells(nFila + 4, j * 2 + 1)
        nTotAcumME = obj_Excel.Cells(nFila + 4, j * 2 + 2)
        If nTotAcumMN < 0 Then
            Set celda = obj_Excel.Cells(nFila, j * 2 + 1)
            celda.value = Abs(nTotAcumMN) + 500
        End If
        If nTotAcumME < 0 Then
            Set celda = obj_Excel.Cells(nFila, j * 2 + 2)
            celda.value = Abs(nTotAcumME) + 150
        End If
    Next
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeTop).Weight = xlThin
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeBottom).Weight = xlThin
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeLeft).Weight = xlThin
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlEdgeRight).Weight = xlThin
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlInsideVertical).Weight = xlThin
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    obj_Excel.Range("B39:N" & nFila + 1).Borders(xlInsideHorizontal).Weight = xlThin
    obj_Excel.Range("B" & nFila + 1 & ":N" & nFila + 1).Font.Bold = True
    obj_Excel.Range("B" & nFila + 1 & ":N" & nFila + 1).Cells.Interior.Color = RGB(255, 255, 153)
    

    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeTop).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeBottom).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeLeft).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlEdgeRight).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlInsideVertical).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlInsideVertical).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Borders(xlInsideHorizontal).Weight = xlThin
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Font.Bold = True
    obj_Excel.Range("B" & nFila + 3 & ":N" & nFila + 5).Cells.Interior.Color = RGB(255, 255, 153)

End Sub
Private Sub EstableceFormulas(ByVal obj_Excel As Excel.Application, ByVal nFila As Integer, ByVal sConcepto As String, ByVal sFormula As String)
        Dim celda As Excel.Range
        
        Set celda = obj_Excel.Range("B" & nFila)
        celda.value = Replace(sConcepto, "[X]", "B")
        
        Set celda = obj_Excel.Range("C" & nFila)
        celda.value = Replace(sFormula, "[X]", "C")
        
        Set celda = obj_Excel.Range("D" & nFila)
        celda.value = Replace(sFormula, "[X]", "D")
        Set celda = obj_Excel.Range("E" & nFila)
        celda.value = Replace(sFormula, "[X]", "E")
        Set celda = obj_Excel.Range("F" & nFila)
        celda.value = Replace(sFormula, "[X]", "F")
        Set celda = obj_Excel.Range("G" & nFila)
        celda.value = Replace(sFormula, "[X]", "G")
        Set celda = obj_Excel.Range("H" & nFila)
        celda.value = Replace(sFormula, "[X]", "H")
        Set celda = obj_Excel.Range("I" & nFila)
        celda.value = Replace(sFormula, "[X]", "I")
        Set celda = obj_Excel.Range("J" & nFila)
        celda.value = Replace(sFormula, "[X]", "J")
        Set celda = obj_Excel.Range("K" & nFila)
        celda.value = Replace(sFormula, "[X]", "K")
        Set celda = obj_Excel.Range("L" & nFila)
        celda.value = Replace(sFormula, "[X]", "L")
        Set celda = obj_Excel.Range("M" & nFila)
        celda.value = Replace("=C[X]+E[X]+G[X]+I[X]+K[X]", "[X]", CStr(nFila))
        Set celda = obj_Excel.Range("N" & nFila)
        celda.value = Replace("=D[X]+F[X]+H[X]+J[X]+L[X]", "[X]", CStr(nFila))
End Sub
Private Sub ActivoCreditos(ByVal ldFecha As Date, ByVal obj_Excel As Excel.Application)
    Dim oCtaIf As NCajaCtaIF
    Set oCtaIf = New NCajaCtaIF
    Dim oCtaCont As dBalanceCont
    Set oCtaCont = New dBalanceCont
    Dim prsVig As ADODB.Recordset
    Dim prsRef As ADODB.Recordset
    Dim celda As Excel.Range
    
    Dim nCredVig1MN As Double: Dim nCredVig1ME As Double
    Dim nCredVig2MN As Double: Dim nCredVig2ME As Double
    Dim nCredVig3MN As Double: Dim nCredVig3ME As Double
    Dim nCredVig4_6MN As Double: Dim nCredVig4_6ME As Double
    Dim nCredVig7_12MN As Double: Dim nCredVig7_12ME As Double
    Dim nCredVigMas12MN As Double: Dim nCredVigMas12ME As Double
   
    Dim nCredRef1MN As Double: Dim nCredRef1ME As Double
    Dim nCredRef2MN As Double: Dim nCredRef2ME As Double
    Dim nCredRef3MN As Double: Dim nCredRef3Me As Double
    Dim nCredRef4_6MN As Double: Dim nCredRef4_6ME As Double
    Dim nCredRef7_12MN As Double: Dim nCredRef7_12ME As Double
    Dim nCredRefMas12MN As Double: Dim nCredRefMas12ME As Double
    
    Dim nMoraMN As Double
    Dim nMoraME As Double
    Dim nPorcMora As Double
    
    Set prsVig = oCtaIf.GetCreditosVigentes(Format(ldFecha, "yyyymmdd"))
    nCredVig1MN = prsVig!Venc07 + prsVig!Venc08_15 + prsVig!Venc016_30
    nCredVig2MN = prsVig!Venc031_60
    nCredVig3MN = prsVig!Venc061_90
    nCredVig4_6MN = prsVig!Venc091_180
    nCredVig7_12MN = Round(prsVig!Venc0181_360, 2)
    nCredVigMas12MN = Round(prsVig!Venc02_A, 2) + Round(prsVig!Venc03_A, 2) + Round(prsVig!Venc04_A, 2) + Round(prsVig!Venc05_A, 2) + Round(prsVig!Venc06_10A, 2) + Round(prsVig!Venc11_20A, 2) + Round(prsVig!Venc21_Amas, 2)
    
    prsVig.MoveNext
    nCredVig1ME = prsVig!Venc07 + prsVig!Venc08_15 + prsVig!Venc016_30
    nCredVig2ME = prsVig!Venc031_60
    nCredVig3ME = prsVig!Venc061_90
    nCredVig4_6ME = prsVig!Venc091_180
    nCredVig7_12ME = Round(prsVig!Venc0181_360)
    nCredVigMas12ME = Round(prsVig!Venc02_A, 2) + Round(prsVig!Venc03_A, 2) + Round(prsVig!Venc04_A, 2) + Round(prsVig!Venc05_A, 2) + Round(prsVig!Venc06_10A, 2) + Round(prsVig!Venc11_20A, 2) + Round(prsVig!Venc21_Amas, 2)
    
    Set prsRef = oCtaIf.GetCreditosRefinanciados(Format(ldFecha, "yyyymmdd"))
    nCredRef1MN = prsRef!Venc07 + prsRef!Venc08_15 + prsRef!Venc016_30 + oCtaCont.ObtenerCtaContBalanceMensual("1418", ldFecha, "1", Me.txtTipCambio.Text)
    nCredRef2MN = prsRef!Venc031_60
    nCredRef3MN = prsRef!Venc061_90
    nCredRef4_6MN = prsRef!Venc091_180
    nCredRef7_12MN = Round(prsRef!Venc0181_360, 2)
    nCredRefMas12MN = Round(prsRef!Venc02_A, 2) + Round(prsRef!Venc03_A, 2) + Round(prsRef!Venc04_A, 2) + Round(prsRef!Venc05_A, 2) + Round(prsRef!Venc06_10A, 2) + Round(prsRef!Venc11_20A, 2) + Round(prsRef!Venc21_Amas, 2)
    
    prsRef.MoveNext
    nCredRef1ME = prsRef!Venc07 + prsRef!Venc08_15 + prsRef!Venc016_30 + oCtaCont.ObtenerCtaContBalanceMensual("1428", ldFecha, "2", Me.txtTipCambio.Text)
    nCredRef2ME = prsRef!Venc031_60
    nCredRef3Me = prsRef!Venc061_90
    nCredRef4_6ME = prsRef!Venc091_180
    nCredRef7_12ME = prsRef!Venc0181_360
    nCredRefMas12ME = Round(prsRef!Venc02_A, 2) + Round(prsRef!Venc03_A, 2) + Round(prsRef!Venc04_A, 2) + Round(prsRef!Venc05_A, 2) + Round(prsRef!Venc06_10A, 2) + Round(prsRef!Venc11_20A, 2) + Round(prsRef!Venc21_Amas, 2)
    
    
    nPorcMora = txtMora.Text / 100
    If sOpeReporte = "770162" Then
        nPorcMora = Round((oCtaCont.ObtenerCtaContBalanceMensual("1404,1405,1406", ldFecha, "0", Me.txtTipCambio.Text) / oCtaCont.ObtenerCtaContBalanceMensual("14", ldFecha, "0", Me.txtTipCambio.Text) + nPorcMora), 2)
    End If
    
    Set celda = obj_Excel.Range("C13")
    celda.value = Round((nCredVig1MN + nCredRef1MN) * (1 - nPorcMora), 2)
    nMoraMN = Round(nMoraMN + (nCredVig1MN + nCredRef1MN) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("D13")
    celda.value = Round((nCredVig1ME + nCredRef1ME) * (1 - nPorcMora), 2)
    nMoraME = Round(nMoraME + (nCredVig1ME + nCredRef1ME) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("E13")
    celda.value = Round((nCredVig2MN + nCredRef2MN) * (1 - nPorcMora), 2)
    nMoraMN = Round(nMoraMN + (nCredVig2MN + nCredRef2MN) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("F13")
    celda.value = Round((nCredVig2ME + nCredRef2ME) * (1 - nPorcMora), 2)
    nMoraME = Round(nMoraME + (nCredVig2ME + nCredRef2ME) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("G13")
    celda.value = Round((nCredVig3MN + nCredRef3MN) * (1 - nPorcMora), 2)
    nMoraMN = Round(nMoraMN + (nCredVig3MN + nCredRef3MN) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("H13")
    celda.value = Round((nCredVig3ME + nCredRef3Me) * (1 - nPorcMora), 2)
    nMoraME = Round(nMoraME + (nCredVig3ME + nCredRef3Me) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("I13")
    celda.value = Round((nCredVig4_6MN + nCredRef4_6MN) * (1 - nPorcMora), 2)
    nMoraMN = Round(nMoraMN + (nCredVig4_6MN + nCredRef4_6MN) * nPorcMora, 2)
    
    Set celda = obj_Excel.Range("J13")
    celda.value = Round((nCredVig4_6ME + nCredRef4_6ME) * (1 - nPorcMora), 2)
    nMoraME = Round(nMoraME + (nCredVig4_6ME + nCredRef4_6ME) * nPorcMora, 2)
    
    Dim tmpMN As Double: Dim tmpME As Double
    If sOpeReporte = "770162" Then
        tmpMN = Round((nCredVig7_12MN + nCredRef7_12MN) * (1 - nPorcMora), 2)
        nMoraMN = Round(nMoraMN + (nCredVig7_12MN + nCredRef7_12MN) * nPorcMora, 2)
        
        tmpME = Round((nCredVig7_12ME + nCredRef7_12ME) * (1 - nPorcMora), 2)
        nMoraME = Round(nMoraME + (nCredVig7_12ME + nCredRef7_12ME) * nPorcMora, 2)
        
        nMoraMN = Round(nMoraMN + (nCredVigMas12MN + nCredRefMas12MN) * nPorcMora, 2)
        Set celda = obj_Excel.Range("K13")
        celda.value = Round((nCredVigMas12MN + nCredRefMas12MN) * (1 - nPorcMora) + tmpMN + nMoraMN, 2)
        
        nMoraME = Round(nMoraME + (nCredVigMas12ME + nCredRefMas12ME) * nPorcMora, 2)
        Set celda = obj_Excel.Range("L13")
        celda.value = Round((nCredVigMas12ME + nCredRefMas12ME) * (1 - nPorcMora) + tmpME + nMoraME, 2)
    Else
        Set celda = obj_Excel.Range("K13")
        celda.value = Round((nCredVig7_12MN + nCredRef7_12MN) * (1 - nPorcMora), 2)
        nMoraMN = Round(nMoraMN + (nCredVig7_12MN + nCredRef7_12MN) * nPorcMora, 2)
        
        Set celda = obj_Excel.Range("L13")
        celda.value = Round((nCredVig7_12ME + nCredRef7_12ME) * (1 - nPorcMora), 2)
        nMoraME = Round(nMoraME + (nCredVig7_12ME + nCredRef7_12ME) * nPorcMora, 2)
        
        nMoraMN = Round(nMoraMN + (nCredVigMas12MN + nCredRefMas12MN) * nPorcMora, 2)
        Set celda = obj_Excel.Range("M13")
        celda.value = Round((nCredVigMas12MN + nCredRefMas12MN) * (1 - nPorcMora) + nMoraMN, 2)
        
        nMoraME = Round(nMoraME + (nCredVigMas12ME + nCredRefMas12ME) * nPorcMora, 2)
        Set celda = obj_Excel.Range("N13")
        celda.value = Round((nCredVigMas12ME + nCredRefMas12ME) * (1 - nPorcMora) + nMoraME, 2)
    End If
    
    Set prsRef = Nothing
    Set prsVig = Nothing
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
End Sub
Private Sub PasivoDepositoPlazo(ByVal ldFecha As Date, ByVal obj_Excel As Excel.Application)
    Dim oCtaIf As NCajaCtaIF
    Dim oCtaCont As dBalanceCont
    Dim prs As ADODB.Recordset
    Dim celda As Excel.Range
    
    Dim nDep1MN As Double: Dim nDep1ME As Double
    Dim nDep2MN As Double: Dim nDep2ME As Double
    Dim nDep3MN As Double: Dim nDep3ME As Double
    Dim nDep4_6MN As Double: Dim nDep4_6ME As Double
    Dim nDep7_12MN As Double: Dim nDep7_12ME As Double
    Dim nDepMas12MN As Double: Dim nDepMas12ME As Double
    
    Set oCtaIf = New NCajaCtaIF
    Set oCtaCont = New dBalanceCont
    
    Set prs = oCtaIf.GetPlazoFijoAnexo7(Format(ldFecha, "yyyymmdd"), Me.txtTipCambio.Text)
    
    nDep1MN = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30 + oCtaCont.ObtenerCtaContBalanceMensual("211305", ldFecha, "1", Me.txtTipCambio.Text)
    nDep2MN = prs!Venc31a60
    nDep3MN = prs!Venc61a90
    nDep4_6MN = prs!Venc91a180
    nDep7_12MN = prs!Venc181a360
    nDepMas12MN = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas
    
    prs.MoveNext
    nDep1ME = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30 + oCtaCont.ObtenerCtaContBalanceMensual("212305", ldFecha, "2", Me.txtTipCambio.Text)
    nDep2ME = prs!Venc31a60
    nDep3ME = prs!Venc61a90
    nDep4_6ME = prs!Venc91a180
    nDep7_12ME = prs!Venc181a360
    nDepMas12ME = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas
    
    Set prs = Nothing
    
    Set prs = oCtaIf.GetPlazoFijosDeOtros(Format(ldFecha, "yyyymmdd"), Me.txtTipCambio.Text)
    
    nDep1MN = nDep1MN - (prs!Venc07 + prs!Venc08a15 + prs!Venc16a30)
    nDep2MN = nDep2MN - prs!Venc31a60
    nDep3MN = nDep3MN - prs!Venc61a90
    nDep4_6MN = nDep4_6MN - prs!Venc91a180
    nDep7_12MN = nDep7_12MN - prs!Venc181a360
    nDepMas12MN = nDepMas12MN - (prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas)
    
    prs.MoveNext
    nDep1ME = nDep1ME - (prs!Venc07 + prs!Venc08a15 + prs!Venc16a30)
    nDep2ME = nDep2ME - prs!Venc31a60
    nDep3ME = nDep3ME - prs!Venc61a90
    nDep4_6ME = nDep4_6ME - prs!Venc91a180
    nDep7_12ME = nDep7_12ME - prs!Venc181a360
    nDepMas12ME = nDepMas12ME - (prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas)
    
    nDep7_12MN = nDep7_12MN - oCtaCont.ObtenerCtaContBalanceMensual("2117", ldFecha, "1", Me.txtTipCambio.Text)
    nDep7_12ME = nDep7_12ME - oCtaCont.ObtenerCtaContBalanceMensual("2127", ldFecha, "2", Me.txtTipCambio.Text)
    
    If sOpeReporte = "770162" Then
        nDep7_12MN = nDep7_12MN + nDepMas12MN
        nDep7_12ME = nDep7_12ME + nDepMas12ME
        
        nDep1MN = nDep1MN + (nDep2MN * 0.5)
        nDep2MN = nDep2MN * 0.5 + nDep3MN * 0.5
        nDep3MN = nDep3MN * 0.5 + nDep4_6MN * 0.5
        nDep4_6MN = nDep4_6MN * 0.5 + nDep7_12MN * 0.5
        nDep7_12MN = nDep7_12MN * 0.5
        
        nDep1ME = nDep1ME + (nDep2ME * 0.5)
        nDep2ME = nDep2ME * 0.5 + nDep3ME * 0.5
        nDep3ME = nDep3ME * 0.5 + nDep4_6ME * 0.5
        nDep4_6ME = nDep4_6ME * 0.5 + nDep7_12ME * 0.5
        nDep7_12ME = nDep7_12ME * 0.5
    End If
    
    Set celda = obj_Excel.Range("C22")
    celda.value = nDep1MN
    Set celda = obj_Excel.Range("D22")
    celda.value = nDep1ME
    Set celda = obj_Excel.Range("E22")
    celda.value = nDep2MN
    Set celda = obj_Excel.Range("F22")
    celda.value = nDep2ME
    Set celda = obj_Excel.Range("G22")
    celda.value = nDep3MN
    Set celda = obj_Excel.Range("H22")
    celda.value = nDep3ME
    Set celda = obj_Excel.Range("I22")
    celda.value = nDep4_6MN
    Set celda = obj_Excel.Range("J22")
    celda.value = nDep4_6ME
    Set celda = obj_Excel.Range("K22")
    celda.value = nDep7_12MN
    Set celda = obj_Excel.Range("L22")
    celda.value = nDep7_12ME
    If sOpeReporte <> "770162" Then
        Set celda = obj_Excel.Range("M22")
        celda.value = nDepMas12MN
        Set celda = obj_Excel.Range("N22")
        celda.value = nDepMas12ME
    End If

    Set prs = Nothing
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
End Sub
Public Sub DepositosIFISyOFI(ByVal ldFecha As Date, ByVal obj_Excel As Excel.Application)
    Dim oCtaIf As NCajaCtaIF
    Dim oCtaCont As dBalanceCont
    Dim prs As ADODB.Recordset
    Dim celda As Excel.Range
    
    Dim nDep1MN As Double: Dim nDep1ME As Double
    Dim nDep2MN As Double: Dim nDep2ME As Double
    Dim nDep3MN As Double: Dim nDep3ME As Double
    Dim nDep4_6MN As Double: Dim nDep4_6ME As Double
    Dim nDep7_12MN As Double: Dim nDep7_12ME As Double
    Dim nDepMas12MN As Double: Dim nDepMas12ME As Double
    Dim nMontoAhorroMN As Double: Dim nMontoAhorroME As Double
    Dim nMes As Integer
    
    Set oCtaCont = New dBalanceCont
    Set oCtaIf = New NCajaCtaIF
    Set prs = oCtaIf.GetPlazoFijosDeOtros(Format(ldFecha, "yyyymmdd"), Me.txtTipCambio.Text)
    
    nDep1MN = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30 + oCtaCont.ObtenerCtaContBalanceMensual("2318", ldFecha, "1", Me.txtTipCambio.Text)
    nDep2MN = prs!Venc31a60
    nDep3MN = prs!Venc61a90
    nDep4_6MN = prs!Venc91a180
    nDep7_12MN = prs!Venc181a360
    nDepMas12MN = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas
    
    prs.MoveNext
    nDep1ME = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30 + oCtaCont.ObtenerCtaContBalanceMensual("2328", ldFecha, "2", Me.txtTipCambio.Text)
    nDep2ME = prs!Venc31a60
    nDep3ME = prs!Venc61a90
    nDep4_6ME = prs!Venc91a180
    nDep7_12ME = prs!Venc181a360
    nDepMas12ME = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas

     nMontoAhorroMN = oCtaCont.ObtenerCtaContBalanceMensual("2312", ldFecha, "1", Me.txtTipCambio.Text)
     nMontoAhorroME = oCtaCont.ObtenerCtaContBalanceMensual("2322", ldFecha, "2", Me.txtTipCambio.Text)
     nMes = CInt(Right(cmbMes.Text, 2))

     nDep1MN = nDep1MN + Abs(DistribucionMontos(nMes + 1, nMontoAhorroMN))
     nDep2MN = nDep2MN + Abs(DistribucionMontos(nMes + 2, nMontoAhorroMN))
     nDep3MN = nDep3MN + Abs(DistribucionMontos(nMes + 3, nMontoAhorroMN))
     nDep4_6MN = nDep4_6MN + Abs(DistribucionMontos(nMes + 4, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 5, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 6, nMontoAhorroMN))
     nDep7_12MN = nDep7_12MN + Abs(DistribucionMontos(nMes + 7, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroMN))
     nDepMas12MN = nMontoAhorroMN
     
     nDep1ME = nDep1ME + Abs(DistribucionMontos(nMes + 1, nMontoAhorroME))
     nDep2ME = nDep2ME + Abs(DistribucionMontos(nMes + 2, nMontoAhorroME))
     nDep3ME = nDep3ME + Abs(DistribucionMontos(nMes + 3, nMontoAhorroME))
     nDep4_6ME = nDep4_6ME + Abs(DistribucionMontos(nMes + 4, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 5, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 6, nMontoAhorroME))
     nDep7_12ME = nDep7_12ME + Abs(DistribucionMontos(nMes + 7, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroME))
     nDepMas12ME = nMontoAhorroME
    
    If sOpeReporte = "770162" Then 'Si es Anexo16B
        Set celda = obj_Excel.Range("C26")
        celda.value = nDep1MN + nDep2MN + nDep3MN + nDep4_6MN + nDep7_12MN + nDepMas12MN
        Set celda = obj_Excel.Range("D26")
        celda.value = nDep1ME + nDep2ME + nDep3ME + nDep4_6ME + nDep7_12ME + nDepMas12ME
    Else
        Set celda = obj_Excel.Range("C26")
        celda.value = Round(nDep1MN, 2)
        Set celda = obj_Excel.Range("D26")
        celda.value = Round(nDep1ME, 2)
        Set celda = obj_Excel.Range("E26")
        celda.value = Round(nDep2MN, 2)
        Set celda = obj_Excel.Range("F26")
        celda.value = Round(nDep2ME, 2)
        Set celda = obj_Excel.Range("G26")
        celda.value = Round(nDep3MN, 2)
        Set celda = obj_Excel.Range("H26")
        celda.value = Round(nDep3ME, 2)
        Set celda = obj_Excel.Range("I26")
        celda.value = Round(nDep4_6MN, 2)
        Set celda = obj_Excel.Range("J26")
        celda.value = Round(nDep4_6ME, 2)
        Set celda = obj_Excel.Range("K26")
        celda.value = Round(nDep7_12MN, 2)
        Set celda = obj_Excel.Range("L26")
        celda.value = Round(nDep7_12ME, 2)
        Set celda = obj_Excel.Range("M26")
        celda.value = Round(nDepMas12MN, 2)
        Set celda = obj_Excel.Range("N26")
        celda.value = Round(nDepMas12ME, 2)
    End If
    
    Set prs = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
End Sub
Private Sub ActivoDisponible(ByVal ldFecha As Date, ByVal obj_Excel As Excel.Application)
    Dim nEncajeMN As Double
    Dim nEncajeME As Double
    Dim nMes As Integer
    Dim celda As Excel.Range
    
        
    Dim oCtaCont As dBalanceCont
    Set oCtaCont = New dBalanceCont
    
        nEncajeMN = CDbl(Me.txtEncajeMN.Text)
        nEncajeME = CDbl(Me.txtEncajeME.Text)
        nMes = CInt(Right(cmbMes.Text, 2))

        
        Set celda = obj_Excel.Range("C10")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("11", ldFecha, "1", Me.txtTipCambio.Text) - nEncajeMN + Abs(DistribucionMontos(nMes + 1, nEncajeMN))
        Set celda = obj_Excel.Range("D10")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("11", ldFecha, "2", Me.txtTipCambio.Text) - nEncajeME + Abs(DistribucionMontos(nMes + 1, nEncajeME))
        Set celda = obj_Excel.Range("E10")
        celda.value = Abs(DistribucionMontos(nMes + 2, nEncajeMN))
        Set celda = obj_Excel.Range("F10")
        celda.value = Abs(DistribucionMontos(nMes + 2, nEncajeME))
        Set celda = obj_Excel.Range("G10")
        celda.value = Abs(DistribucionMontos(nMes + 3, nEncajeMN))
        Set celda = obj_Excel.Range("H10")
        celda.value = Abs(DistribucionMontos(nMes + 3, nEncajeME))
        Set celda = obj_Excel.Range("I10")
        celda.value = Abs(DistribucionMontos(nMes + 4, nEncajeMN)) + Abs(DistribucionMontos(nMes + 5, nEncajeMN)) + Abs(DistribucionMontos(nMes + 6, nEncajeMN))
        Set celda = obj_Excel.Range("J10")
        celda.value = Abs(DistribucionMontos(nMes + 4, nEncajeME)) + Abs(DistribucionMontos(nMes + 5, nEncajeME)) + Abs(DistribucionMontos(nMes + 6, nEncajeME))
        
        If sOpeReporte = "770162" Then
            Set celda = obj_Excel.Range("K10")
            celda.value = Abs(DistribucionMontos(nMes + 7, nEncajeMN)) + Abs(DistribucionMontos(nMes + 8, nEncajeMN)) + Abs(DistribucionMontos(nMes + 9, nEncajeMN)) + Abs(DistribucionMontos(nMes + 10, nEncajeMN)) + Abs(DistribucionMontos(nMes + 11, nEncajeMN)) + Abs(DistribucionMontos(nMes + 12, nEncajeMN)) + nEncajeMN
            Set celda = obj_Excel.Range("L10")
            celda.value = Abs(DistribucionMontos(nMes + 7, nEncajeME)) + Abs(DistribucionMontos(nMes + 8, nEncajeME)) + Abs(DistribucionMontos(nMes + 9, nEncajeME)) + Abs(DistribucionMontos(nMes + 10, nEncajeME)) + Abs(DistribucionMontos(nMes + 11, nEncajeME)) + Abs(DistribucionMontos(nMes + 12, nEncajeME)) + nEncajeME
        Else
            Set celda = obj_Excel.Range("K10")
            celda.value = Abs(DistribucionMontos(nMes + 7, nEncajeMN)) + Abs(DistribucionMontos(nMes + 8, nEncajeMN)) + Abs(DistribucionMontos(nMes + 9, nEncajeMN)) + Abs(DistribucionMontos(nMes + 10, nEncajeMN)) + Abs(DistribucionMontos(nMes + 11, nEncajeMN)) + Abs(DistribucionMontos(nMes + 12, nEncajeMN))
            Set celda = obj_Excel.Range("L10")
            celda.value = Abs(DistribucionMontos(nMes + 7, nEncajeME)) + Abs(DistribucionMontos(nMes + 8, nEncajeME)) + Abs(DistribucionMontos(nMes + 9, nEncajeME)) + Abs(DistribucionMontos(nMes + 10, nEncajeME)) + Abs(DistribucionMontos(nMes + 11, nEncajeME)) + Abs(DistribucionMontos(nMes + 12, nEncajeME))
            Set celda = obj_Excel.Range("M10")
            celda.value = nEncajeMN
            Set celda = obj_Excel.Range("N10")
            celda.value = nEncajeME
        End If

        
        Set oCtaCont = Nothing
End Sub
Private Sub cargarCalecAdeudados(ByVal ldFecha As Date, ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim celda As Excel.Range
    Dim oCtaCont As dBalanceCont
    Dim nVenc1MN As Double: Dim nVenc1ME As Double
    Dim nVenc2MN As Double: Dim nVenc2ME As Double
    Dim nVenc3MN As Double: Dim nVenc3ME As Double
    Dim nVenc4_6MN As Double: Dim nVenc4_6ME As Double
    Dim nVenc7_12MN As Double: Dim nVenc7_12ME As Double
    Dim nVencMas12MN As Double: Dim nVencMas12ME As Double
    Dim nInteresMN As Double: Dim nInteresME As Double
    Dim nMes As Integer
    
    Set oCtaCont = New dBalanceCont
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Set prs = oCtaIf.GetCaleAdeudadosCapIntXTramos(Format(ldFecha, "yyyymmdd"), "1")
    
    If Not prs.EOF Or prs.BOF Then
        nVenc1MN = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30
        nVenc2MN = prs!Venc31a60
        nVenc3MN = prs!Venc61a90
        nVenc4_6MN = prs!Venc91a180
        nVenc7_12MN = prs!Venc181a360
        nVencMas12MN = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas
    End If
    
    Set prs = oCtaIf.GetCaleAdeudadosCapIntXTramos(Format(ldFecha, "yyyymmdd"), "2")
    
    If Not prs.EOF Or prs.BOF Then
        nVenc1ME = prs!Venc07 + prs!Venc08a15 + prs!Venc16a30
        nVenc2ME = prs!Venc31a60
        nVenc3ME = prs!Venc61a90
        nVenc4_6ME = prs!Venc91a180
        nVenc7_12ME = prs!Venc181a360
        nVencMas12ME = prs!Venc02_A + prs!Venc03_A + prs!Venc04_A + prs!Venc05_A + prs!Venc06_10A + prs!Venc11_20A + prs!Venc21_Amas
    End If
    
    
    Set celda = pobj_Excel.Range("C27")
    celda.value = nVenc1MN
    Set celda = pobj_Excel.Range("D27")
    celda.value = nVenc1ME
    Set celda = pobj_Excel.Range("E27")
    celda.value = nVenc2MN
    Set celda = pobj_Excel.Range("F27")
    celda.value = nVenc2ME
    Set celda = pobj_Excel.Range("G27")
    celda.value = nVenc3MN
    Set celda = pobj_Excel.Range("H27")
    celda.value = nVenc3ME
    Set celda = pobj_Excel.Range("I27")
    celda.value = nVenc4_6MN
    Set celda = pobj_Excel.Range("J27")
    celda.value = nVenc4_6ME
    
    If sOpeReporte = "770162" Then
        Set celda = pobj_Excel.Range("K27")
        celda.value = nVenc7_12MN + nVencMas12MN
        Set celda = pobj_Excel.Range("L27")
        celda.value = nVenc7_12ME + nVencMas12ME
    Else
        Set celda = pobj_Excel.Range("K27")
        celda.value = nVenc7_12MN
        Set celda = pobj_Excel.Range("L27")
        celda.value = nVenc7_12ME
        Set celda = pobj_Excel.Range("M27")
        celda.value = nVencMas12MN
        Set celda = pobj_Excel.Range("N27")
        celda.value = nVencMas12ME
    End If


    Set prs = Nothing
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
End Sub
Private Sub PasivoDepositoAhorros(ByVal ldFecha As Date, ByVal obj_Excel As Excel.Application)
    Dim nMontoAhorroMN As Double
    Dim nMontoAhorroME As Double
    Dim nMes As Integer
    Dim celda As Excel.Range
    Dim sEscEstres As Boolean
        
    Dim oCtaCont As dBalanceCont
    Set oCtaCont = New dBalanceCont
    
        nMontoAhorroMN = oCtaCont.ObtenerCtaContBalanceMensual("2112", ldFecha, "1", Me.txtTipCambio.Text)
        nMontoAhorroME = oCtaCont.ObtenerCtaContBalanceMensual("2122", ldFecha, "2", Me.txtTipCambio.Text)
        nMes = CInt(Right(cmbMes.Text, 2))

        If sOpeReporte = "770162" Then
            sEscEstres = True
        End If
        
        Set celda = obj_Excel.Range("C21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 1, nMontoAhorroMN, sEscEstres)), 2)
        Set celda = obj_Excel.Range("D21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 1, nMontoAhorroME, sEscEstres)), 2)
        Set celda = obj_Excel.Range("E21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 2, nMontoAhorroMN)), 2)
        Set celda = obj_Excel.Range("F21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 2, nMontoAhorroME)), 2)
        Set celda = obj_Excel.Range("G21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 3, nMontoAhorroMN)), 2)
        Set celda = obj_Excel.Range("H21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 3, nMontoAhorroME)), 2)
        Set celda = obj_Excel.Range("I21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 4, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 5, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 6, nMontoAhorroMN)), 2)
        Set celda = obj_Excel.Range("J21")
        celda.value = Round(Abs(DistribucionMontos(nMes + 4, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 5, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 6, nMontoAhorroME)), 2)
        
        If sOpeReporte = "770162" Then
            Set celda = obj_Excel.Range("K21")
            celda.value = Round(Abs(DistribucionMontos(nMes + 7, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroMN)) + nMontoAhorroMN, 2)
            Set celda = obj_Excel.Range("L21")
            celda.value = Round(Abs(DistribucionMontos(nMes + 7, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroME)) + nMontoAhorroME, 2)
        Else
            Set celda = obj_Excel.Range("K21")
            celda.value = Round(Abs(DistribucionMontos(nMes + 7, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroMN)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroMN)), 2)
            Set celda = obj_Excel.Range("L21")
            celda.value = Round(Abs(DistribucionMontos(nMes + 7, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 8, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 9, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 10, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 11, nMontoAhorroME)) + Abs(DistribucionMontos(nMes + 12, nMontoAhorroME)), 2)
            Set celda = obj_Excel.Range("M21")
            celda.value = Round(nMontoAhorroMN, 2)
            Set celda = obj_Excel.Range("N21")
            celda.value = Round(nMontoAhorroME, 2)
        End If

        
        Set oCtaCont = Nothing
End Sub
Private Function DistribucionMontos(ByVal MES As Integer, ByRef pnMonto As Double, Optional psEscenEstres As Boolean = False) As Double
    Dim nMontoBanda As Double
    If psEscenEstres = False Then
        If MES > 12 Then
            MES = MES - 12
        End If
        Select Case MES
            Case 1:
                nMontoBanda = pnMonto * (-10 / 100)
            Case 2:
                nMontoBanda = pnMonto * (-6 / 100)
            Case 3:
                nMontoBanda = pnMonto * (-1 / 100)
            Case 4:
                nMontoBanda = pnMonto * (0 / 100)
            Case 5:
                nMontoBanda = pnMonto * (-2 / 100)
            Case 6:
                nMontoBanda = pnMonto * (-3 / 100)
            Case 7:
                nMontoBanda = pnMonto * (-1 / 100)
            Case 8:
                nMontoBanda = pnMonto * (-2 / 100)
            Case 9:
                nMontoBanda = pnMonto * (0 / 100)
            Case 10:
                nMontoBanda = pnMonto * (0 / 100)
            Case 11:
                nMontoBanda = pnMonto * (-5 / 100)
            Case 12:
                nMontoBanda = pnMonto * (-8 / 100)
        End Select
    Else
        nMontoBanda = pnMonto * (-30 / 100)
    End If
    
    pnMonto = pnMonto + nMontoBanda
    DistribucionMontos = nMontoBanda
End Function
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function

Private Sub cmdQuitar_Click()
    Me.grdContingencia.EliminaFila Me.grdContingencia.Row
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
    Me.cmbMes.ListIndex = Month(gdFecSis) - 1
    txtTipCambio.Text = TipoCambioCierre(mskAnio.Text, cmbMes.ListIndex + 1)
End Sub

Public Sub Inicio(psOpeReporte As String, sTitForm As String)
    sOpeReporte = psOpeReporte
    Me.Caption = sTitForm
    If sOpeReporte = "770160" Then
        cmdGenerar.Top = 840
        cmdSalir.Top = 840
        fraPC.Visible = False
        Me.Height = 2000
        'Me.chkPlanC.Visible = False
    Else
        'Call chkPlanC_Click
    End If
    Me.Show 1
End Sub


Private Sub mskAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTipCambio.Text = TipoCambioCierre(mskAnio, cmbMes.ListIndex + 1)
        txtTipCambio.SetFocus
    End If
End Sub

Private Sub txtEncajeME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If sOpeReporte = "770162" Then
            Me.cmdAgregar.SetFocus
        Else
            Me.cmdGenerar.SetFocus
        End If
    End If
End Sub

Private Sub txtEncajeMN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtEncajeME.SetFocus
    End If
End Sub
Private Sub txtMora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPatEfec.SetFocus
    End If
End Sub

Private Sub txtPatEfec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtEncajeMN.SetFocus
    End If
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMora.SetFocus
    End If
End Sub
