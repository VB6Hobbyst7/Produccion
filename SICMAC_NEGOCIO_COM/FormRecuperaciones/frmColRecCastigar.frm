VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColRecCastigar 
   Caption         =   "Recuperaciones - Castigar Crédito"
   ClientHeight    =   6675
   ClientLeft      =   1965
   ClientTop       =   2265
   ClientWidth     =   8910
   Icon            =   "frmColRecCastigar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8910
   Begin VB.Frame FraSeleccion 
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   3555
      Begin VB.OptionButton OpSeleccion 
         Caption         =   "Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OpSeleccion 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdArchivo 
      Caption         =   "&Cargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   435
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "108"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdCastigar 
      Caption         =   "C&astigar"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   6240
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
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
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   3915
      Begin VB.OptionButton Opt 
         Caption         =   "Excel"
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Cuenta"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         Caption         =   "General"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lstJudicial 
      Height          =   5445
      Left            =   60
      TabIndex        =   5
      Top             =   720
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9604
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   3660
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Archivo de Excel (*.xls)|*.xls"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "frmColRecCastigar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* CASTIGAR CREDITO EN RECUPERACIONES
'Archivo:  frmColRecCastigar.frm
'LAYG   :  15/08/2001.
'Resumen:  Nos permite realizar el Castigo de cred en Recuperaciones

Option Explicit
Option Base 1
Dim fsServerConsol As String
Dim sFilename As String

'**DAOR 20070827 *******************************
Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer
'0 --> No Calcula
'1 --> Capital
'2 --> Capital + Int Comp
'3 --> Capital + Int comp + Int Morat
Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
'0 INTERES SIMPLE
'1 INTERES COMPUESTO
'***********************************************
Private Sub cmdArchivo_Click()
    CargaArchivoCastigar sFilename
    FraSeleccion.Visible = True
End Sub

Private Sub cmdBuscar_Click()
Dim lrBusca As ADODB.Recordset
Dim loValCreditos As COMNColocRec.NColRecValida
Dim lsmensaje As String

On Error GoTo ControlError
    PB.Visible = True
    Set lrBusca = New ADODB.Recordset
    Set loValCreditos = New COMNColocRec.NColRecValida
        If Me.Opt(0).value = True Then ' Busca por Codigo
            Set lrBusca = loValCreditos.nBuscaCreditoParaCastigar(fsServerConsol, AXCodCta.NroCuenta, , gsCodAge, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If
        ElseIf Me.Opt(1).value = True Then ' General
            Set lrBusca = loValCreditos.nBuscaCreditoParaCastigar(fsServerConsol, , , gsCodAge, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If
        End If
    Set loValCreditos = Nothing
    
    If lrBusca Is Nothing Then ' Hubo un Error
        Set lrBusca = Nothing
        Exit Sub
    End If
    If lrBusca.BOF And lrBusca.EOF Then
        MsgBox "No Existen Creditos para Castigar ", vbInformation, "Aviso"
        Exit Sub
    End If

    Call LLenaLista(lrBusca)
    Set lrBusca = Nothing
    PB.Visible = False
     
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub LLenaLista(myRs As ADODB.Recordset)
Dim litmX As ListItem
Dim loCalcula As New COMNColocRec.NCOMColRecCalculos 'NColRecCalculos
Dim lnIntComGen As Double
Dim lnDiasultrans As Integer
'**DAOR 20070827*********
Dim nTasaIntComp As Double, nTasaIntMora As Double
Dim nDiasUltTrans As Integer
Dim sFecUltPago As String
Dim nIntCompGenerado As Currency, nIntMoraGenerado As Currency
Dim nSaldoCap As Currency, nSaldoIntComp As Currency, nSaldoIntMora As Currency, nSaldoGasto As Currency
'************************

Set loCalcula = New COMNColocRec.NCOMColRecCalculos 'DAOR 20070827

lstJudicial.ListItems.Clear
Do While Not myRs.EOF
    '**DAOR 20070827, Actualizar saldos a la fecha*********************
    nSaldoCap = myRs!nSaldo
    nSaldoIntComp = myRs!nSaldoIntComp
    nSaldoIntMora = myRs!nSaldoIntMor
    nSaldoGasto = myRs!nSaldoGasto
    sFecUltPago = CDate(fgFechaHoraGrab(myRs!cUltimaActualizacion))
    nTasaIntComp = myRs!nTasaIntComp
    nTasaIntMora = myRs!nTasaIntMora
    nDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(sFecUltPago, "dd/mm/yyyy"))

    'Calcula el Int Comp Generado
    If fnTipoCalcIntComp = 0 Then
        nIntCompGenerado = 0
    ElseIf fnTipoCalcIntComp = 1 Then
        If fnFormaCalcIntComp = 1 Then
            nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(nDiasUltTrans, nTasaIntComp, nSaldoCap)
        Else
            nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(nDiasUltTrans, nTasaIntComp, nSaldoCap)
        End If
    ElseIf fnTipoCalcIntComp = 2 Then
        If fnFormaCalcIntComp = 1 Then
            nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(nDiasUltTrans, nTasaIntComp, nSaldoCap + nSaldoIntComp)
        Else
            nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(nDiasUltTrans, nTasaIntComp, nSaldoCap + nSaldoIntComp)
        End If
    ElseIf fnTipoCalcIntComp = 3 Then
        If fnFormaCalcIntComp = 1 Then
            nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(nDiasUltTrans, nTasaIntComp, nSaldoCap + nSaldoIntComp + nSaldoIntMora)
        Else
            nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(nDiasUltTrans, nTasaIntComp, nSaldoCap + nSaldoIntComp + nSaldoIntMora)
        End If
    End If
    'Calcula el Int Mora Generado
    If fnTipoCalcIntMora = 0 Then
        nIntMoraGenerado = 0
    ElseIf fnTipoCalcIntMora = 1 Then
        If fnFormaCalcIntMora = 1 Then
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(nDiasUltTrans, nTasaIntMora, nSaldoCap)
        Else
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(nDiasUltTrans, nTasaIntMora, nSaldoCap)
        End If
    ElseIf fnTipoCalcIntMora = 2 Then
        If fnFormaCalcIntMora = 1 Then
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(nDiasUltTrans, nTasaIntMora, nSaldoCap + nSaldoIntComp)
        Else
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(nDiasUltTrans, nTasaIntMora, nSaldoCap + nSaldoIntComp)
        End If
    ElseIf fnTipoCalcIntMora = 3 Then
        If fnFormaCalcIntMora = 1 Then
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(nDiasUltTrans, nTasaIntMora, nSaldoCap + nSaldoIntComp + nSaldoIntMora)
        Else
            nIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(nDiasUltTrans, nTasaIntMora, nSaldoCap + nSaldoIntComp + nSaldoIntMora)
        End If
    End If
    nSaldoIntComp = myRs!nSaldoIntComp + nIntCompGenerado
    nSaldoIntMora = myRs!nSaldoIntMor + nIntMoraGenerado
    '******************************************************************
    Set litmX = lstJudicial.ListItems.Add(, , myRs!cCtaCod)     'Nro de Cred Pig
        litmX.SubItems(1) = Mid(myRs!cNomCliE, 1, 30)           'Cliente
        litmX.SubItems(2) = Format(myRs!nMontoCol, "#0.00")     'nMontoCol
        litmX.SubItems(3) = Format(nSaldoCap, "#0.00")          'Saldo Capital
        litmX.SubItems(4) = Format(nSaldoIntComp, "#0.00")      'Saldo Int comp
        litmX.SubItems(5) = Format(nSaldoIntMora, "#0.00")      'Saldo Int Mor
        litmX.SubItems(6) = Format(nSaldoGasto, "#0.00")        'Saldo Gasto
        litmX.SubItems(7) = Format(nIntCompGenerado, "#0.00")  'Int Comp Generado
        litmX.SubItems(8) = myRs!nDemanda       'Estado
        litmX.SubItems(9) = Format(myRs!nProvision, "#0.00")    'Provision (Ultima)
        litmX.SubItems(10) = myRs!cAgeCod '20110701
    myRs.MoveNext
Loop
Set loCalcula = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lstJudicial.ListItems.Clear
    Me.AXCodCta.Texto = ""
End Sub

Private Sub cmdCastigar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito 'NColRecCredito

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim rs As ADODB.Recordset
Dim lnItem As Integer
Dim lsCtaCod As String
Dim lnSaldoCap As Double, lnSaldoIntCom As Double, lnSaldoIntMor As Double
Dim lnSaldoGasto As Double, lnIntComGen As Double
Dim lnDemanda As Integer, lnProvision As Double, lcCodAge As Integer
Dim i As Integer
Dim c As Integer
'On Error GoTo ControlError
Set rs = New ADODB.Recordset

c = 0
For i = 1 To Me.lstJudicial.ListItems.Count
    If Me.lstJudicial.ListItems.iTem(i).Checked = True Then
        c = c + 1
    End If
Next

If c = 0 Then
    MsgBox "Seleccion un Credito para Castigar", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox(" Grabar Castigo de Creditos Seleccionados ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
'MADM 20110513 - With
    With rs
        'Crear RecordSet
        .Fields.Append "cCtaCod", adVarChar, 20
        .Fields.Append "nMoneda", adVarChar, 8
        .Fields.Append "nNombre", adVarChar, 35
        .Fields.Append "nSaldoCap", adDouble
        .Fields.Append "nSaldoIntCom", adDouble
        .Fields.Append "nSaldoIntMor", adDouble
        .Fields.Append "nSaldoGasto", adDouble
        .Fields.Append "nIntComGen", adDouble
        .Fields.Append "nDemanda", adInteger
        .Fields.Append "cCodAge", adInteger
        .Open
        'Llenar Recordset
        
    For lnItem = 1 To Me.lstJudicial.ListItems.Count
        If lstJudicial.ListItems(lnItem).Checked = True Then
        
            'Genera el Mov Nro
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
                lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set loContFunct = Nothing
            
            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
            Set loGrabar = New COMNColocRec.NCOMColRecCredito
                lsCtaCod = lstJudicial.ListItems(lnItem)
                lnSaldoCap = lstJudicial.ListItems(lnItem).ListSubItems(3)
                lnSaldoIntCom = lstJudicial.ListItems(lnItem).ListSubItems(4)
                lnSaldoIntMor = lstJudicial.ListItems(lnItem).ListSubItems(5)
                lnSaldoGasto = lstJudicial.ListItems(lnItem).ListSubItems(6)
                lnIntComGen = lstJudicial.ListItems(lnItem).ListSubItems(7)
                lnDemanda = lstJudicial.ListItems(lnItem).ListSubItems(8)
                lnProvision = IIf(lstJudicial.ListItems(lnItem).ListSubItems(9) = "", 0, lstJudicial.ListItems(lnItem).ListSubItems(9))
                lcCodAge = lstJudicial.ListItems(lnItem).ListSubItems(10)
                
                'Al recordset
                 .AddNew
                .Fields("cCtaCod") = lsCtaCod
                .Fields("nMoneda") = IIf(Mid(lsCtaCod, 9, 1) = "1", "SOLES", "DOLARES")
                .Fields("nNombre") = lstJudicial.ListItems(lnItem).ListSubItems(1)
                .Fields("nSaldoCap") = lnSaldoCap
                .Fields("nSaldoIntCom") = lnSaldoIntCom
                .Fields("nSaldoIntMor") = lnSaldoIntMor
                .Fields("nSaldoGasto") = lnSaldoGasto
                .Fields("nIntComGen") = lnIntComGen
                .Fields("nDemanda") = lnDemanda
                .Fields("cCodAge") = lcCodAge
                
                Call loGrabar.nCastigaCredRecup(lsCtaCod, "GCIM", lsFechaHoraGrab, _
                     lsMovNro, lnSaldoCap, lnSaldoIntCom, lnSaldoIntMor, lnSaldoGasto, _
                     lnIntComGen, lnDemanda, lnProvision, False)
            Set loGrabar = Nothing
        End If
    Next lnItem
    End With
    MsgBox " Se ha realizado el Castigo de Creditos seleccionados ", vbInformation, " Aviso "
    Call Impresion(rs)
    sFilename = ""
    Me.Opt(2).value = False
    lstJudicial.ListItems.Clear
    Set rs = Nothing
    
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Sub CargaCabecera()
    lstJudicial.ColumnHeaders.Add , , "NroCuenta", 1950
    lstJudicial.ColumnHeaders.Add , , "Cliente ", 3500, lvwColumnLeft
    lstJudicial.ColumnHeaders.Add , , "Monto Prestado", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Saldo Capital", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Saldo Int.Compen", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Saldo Int.Morat", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Gastos", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Int Generado", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "Estado", 500, lvwColumnLeft
    lstJudicial.ColumnHeaders.Add , , "Provision", 1200, lvwColumnRight
    lstJudicial.ColumnHeaders.Add , , "CodAge", 400, lvwColumnRight
    lstJudicial.View = lvwReport
End Sub

Sub Marcar(valor As Boolean)
Dim i As Integer
 For i = 1 To lstJudicial.ListItems.Count
     lstJudicial.ListItems(i).Checked = valor
 Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 And AXCodCta.Visible = True Then  'F12
        Dim sCuenta As String
        Dim nProducto As Producto
        sCuenta = frmValTarCodAnt.inicia(nProducto, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
    
End Sub

Private Sub Form_Load()
Dim loConstS As COMDConstSistema.NCOMConstSistema
    
    CargaCabecera
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    AXCodCta.NroCuenta = gsCodCMAC & gsCodAge
    
    Set loConstS = New COMDConstSistema.NCOMConstSistema
        fsServerConsol = loConstS.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstS = Nothing
    'MADM 20110502 - COMENTADO X
'    If gsProyectoActual = "H" Then
'        Opt(2).Visible = True
'        CmdArchivo.Visible = True
'    Else
'        Opt(2).Visible = False
'        CmdArchivo.Visible = False
'    End If
    
    '**DAOR 20070828
    CargaParametros
    
End Sub

Private Sub lstJudicial_ItemCheck(ByVal iTem As MSComctlLib.ListItem)
If val(iTem.SubItems(8)) <> 1 And val(iTem.SubItems(8)) <> 2 Then
    MsgBox "Credito no posee definido el tipo de demanda", vbInformation, "aviso"
End If
End Sub

Private Sub OpSeleccion_Click(Index As Integer)
    Dim i As Integer
    If OpSeleccion(0).value = True Then
        If lstJudicial.ListItems.Count > 0 Then
            For i = 1 To lstJudicial.ListItems.Count
                lstJudicial.ListItems.iTem(i).Checked = True
            Next
        End If
    Else
        If lstJudicial.ListItems.Count > 0 Then
            For i = 1 To lstJudicial.ListItems.Count
                lstJudicial.ListItems.iTem(i).Checked = False
            Next
        End If
    End If
End Sub

Private Sub opt_Click(Index As Integer)
Select Case Index
   Case 0
        AXCodCta.Visible = True
        AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
        cmdBuscar.Visible = True
        CmdArchivo.Visible = False
        lstJudicial.ListItems.Clear
        If Me.FraSeleccion.Visible Then
            Me.FraSeleccion.Visible = False
        End If
   Case 1
        AXCodCta.Visible = False
        cmdBuscar.Visible = True
        CmdArchivo.Visible = False
        lstJudicial.ListItems.Clear
        If Me.FraSeleccion.Visible Then
            Me.FraSeleccion.Visible = False
        End If
   Case 2
        cmdBuscar.Visible = False
        lstJudicial.ListItems.Clear
        CmdArchivo.Visible = True
        AXCodCta.Visible = False
        CommonDialog1.ShowOpen
        sFilename = CommonDialog1.Filename
End Select
End Sub

Public Sub CargaArchivoCastigar(ByVal psNomArchivo As String)
    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim varMatriz As Variant
    Dim cNombreHoja As String
    Dim i As Long, n As Long
    Dim nIntCompGenerado As Currency 'MADM 20110504
    Dim loCalcula As New COMNColocRec.NCOMColRecCalculos 'NColRecCalculos
    
    Set xlApp = New Excel.Application
    nIntCompGenerado = 0
    
    If Trim(psNomArchivo) = "" Then
        MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If Trim(psNomArchivo) <> "" Then
    
        Set xlLibro = xlApp.Workbooks.Open(psNomArchivo, True, True, , "")
        cNombreHoja = "Hoja1"
        'validar nombre de hoja
        Set xlHoja = xlApp.Worksheets(cNombreHoja)
        varMatriz = xlHoja.Range("A2:AA2000").value
        xlLibro.Close SaveChanges:=False
        xlApp.Quit
        Set xlHoja = Nothing
        Set xlLibro = Nothing
        Set xlApp = Nothing
        n = UBound(varMatriz)
        Dim litmX As ListItem
        lstJudicial.ListItems.Clear

        For i = 8 To n 'MADM 20110503 2
            If varMatriz(i, 1) = "" Then
                If i = 8 Then
                    MsgBox "Archivo No tiene Estructura Correcta, la informacion debe estar en la Celda (A9 - Y9)", vbCritical, "Mensaje"
                    Exit For
                Else
                    Opt(2).value = False
                    psNomArchivo = ""
                    Exit For
                End If
            Else
                If varMatriz(i, 26) <> "" Then
                    MsgBox "Archivo No tiene Estructura Correcta, Informacion a partir de la Celda (A9 - Y9)", vbCritical, "Mensaje"
                    Exit For
                Else
                    Set litmX = lstJudicial.ListItems.Add(, , varMatriz(i, 2))          'Nro de Cred Pig
                    litmX.SubItems(1) = Mid(varMatriz(i, 7), 1, 30)                'Cliente
                    litmX.SubItems(2) = Format(varMatriz(i, 13), "#0.00")             'nMontoCol
                    litmX.SubItems(3) = Format(varMatriz(i, 15), "#0.00")             'Saldo Capital
                    litmX.SubItems(4) = Format(varMatriz(i, 16), "#0.00")      'Saldo Int comp
                    litmX.SubItems(5) = Format(varMatriz(i, 17), "#0.00")       'Saldo Int Mor
                    litmX.SubItems(6) = Format(varMatriz(i, 18), "#0.00")        'Saldo Gasto
                    
                    'litmX.SubItems(7) = Format(varMatriz(i, 8), "#0.00")         'Int Comp Generado
                    'Calcula el Int Comp Generado
                        If fnTipoCalcIntComp = 0 Then
                            nIntCompGenerado = 0
                        ElseIf fnTipoCalcIntComp = 1 Then
                            If fnFormaCalcIntComp = 1 Then
                                nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15))
                            Else
                                nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15))
                            End If
                        ElseIf fnTipoCalcIntComp = 2 Then
                            If fnFormaCalcIntComp = 1 Then
                                nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15) + varMatriz(i, 16))
                            Else
                                nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15) + varMatriz(i, 16))
                            End If
                        ElseIf fnTipoCalcIntComp = 3 Then
                            If fnFormaCalcIntComp = 1 Then
                                nIntCompGenerado = loCalcula.nCalculaIntCompGenerado(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15) + varMatriz(i, 16) + varMatriz(i, 17))
                            Else
                                nIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(varMatriz(i, 24), varMatriz(i, 23), varMatriz(i, 15) + varMatriz(i, 16) + varMatriz(i, 17))
                            End If
                        End If
                    litmX.SubItems(7) = Format(nIntCompGenerado, "#0.00")
                    litmX.SubItems(8) = varMatriz(i, 22)      'Estado
                    litmX.SubItems(9) = Format(varMatriz(i, 14), "#0.00")    'Provision (Ultima)
                    litmX.SubItems(10) = varMatriz(i, 25)   'CODAGE
                    litmX.Checked = True
                    Me.OpSeleccion(0).value = True
                End If
            End If
        Next i
    End If

End Sub

'**DAOR 20070827
Private Sub CargaParametros()
Dim loParam As COMDConstSistema.NCOMConstSistema
    Set loParam = New COMDConstSistema.NCOMConstSistema
        fnTipoCalcIntComp = loParam.LeeConstSistema(151)
        fnTipoCalcIntMora = loParam.LeeConstSistema(152)
        fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
        fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA
    Set loParam = Nothing
End Sub
''MADM 20110513
Private Sub Impresion(ByVal pRs As ADODB.Recordset)
    Dim oPrev As previo.clsprevio
    Dim sCad As String
    Dim loImpre As COMNColocRec.NCOMColRecImpre

    Set loImpre = New COMNColocRec.NCOMColRecImpre
    sCad = loImpre.ImpresionCastigo(pRs, gsNomCmac, gdFecSis, gsNomAge, gsCodUser)
    Set loImpre = Nothing

    Set oPrev = New previo.clsprevio
    oPrev.Show sCad, "Pase a Castigo"
    Set oPrev = Nothing
End Sub
