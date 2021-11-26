VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCredVinculados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vinculacion Crediticia"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmCredVinculados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   2970
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   3105
      TabIndex        =   3
      Top             =   2895
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4155
      TabIndex        =   2
      Top             =   2895
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   690
      Left            =   1875
      TabIndex        =   1
      Top             =   3255
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   1217
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   5100
      Begin VB.CheckBox chkUltFinMes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Ultimo fin de mes"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   2535
         Width           =   2115
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2865
         TabIndex        =   12
         Top             =   225
         Width           =   1875
      End
      Begin VB.ListBox lstAge 
         Appearance      =   0  'Flat
         Height          =   2280
         Left            =   2865
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtporTot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "7.00"
         Top             =   840
         Width           =   1545
      End
      Begin VB.TextBox txtPorpar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "5.00"
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "Parcial (%)"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblPocentajeTot 
         Caption         =   "Total (%)"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblMonto 
         Caption         =   "Patrimonio Efectivo :"
         Height          =   480
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   795
      End
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmCredVinculados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim lbSoloTitulares As Boolean
Dim lsCaption As String
'Private fgFecActual As Date 'WIOR 20140128'ORCR20140314

Public Sub Ini(pbSoloTitulares As Boolean, psCaption As String)
    lbSoloTitulares = pbSoloTitulares
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Sub ChkTodas_Click()
    Dim lnI As Integer
    
    For lnI = 0 To Me.lstAge.ListCount - 1
        If Me.chkTodas.value = 1 Then
            lstAge.Selected(lnI) = True
        Else
            lstAge.Selected(lnI) = False
        End If
    Next lnI
End Sub

Private Sub CmdProcesar_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnI As Integer
    Dim lnJ As Integer
    Dim lsCadPers As String
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lsCodAnt As String
    Dim lnAcumulador As Currency
    Dim lnPosAnt As Integer
    Dim clsRep As DCaptaReportes
    Set clsRep = New DCaptaReportes
    
    GetTipCambio gdFecSis
    
    If Not IsNumeric(Me.txtMonto.Text) Then
        MsgBox "Debe ingresar un monto Valido."
        Me.txtMonto.SetFocus
        Exit Sub
    ElseIf CCur(Me.txtMonto.Text) = 0 Then
        MsgBox "Debe ingresar un monto Valido."
        Me.txtMonto.SetFocus
        Exit Sub
    End If
    
    
    Set rs = clsRep.GetFamiliares
    
'    Flex.Cols = 13
'    Flex.Rows = 1
'    Me.Flex.TextMatrix(0, 0) = "Cod Empleado"
'    Me.Flex.TextMatrix(0, 1) = "Nombre"
'    Me.Flex.TextMatrix(0, 2) = "CodigoPersona"
'    Me.Flex.TextMatrix(0, 3) = "Nombre Fam"
'    Me.Flex.TextMatrix(0, 4) = "Relacion Familiar"
'    Me.Flex.TextMatrix(0, 5) = "Monto Creditos Otorgados"
'    Me.Flex.TextMatrix(0, 6) = "Saldos de Creditos"
'    Me.Flex.TextMatrix(0, 7) = "Monto Consolidado por Grupo Familiar"
'    Me.Flex.TextMatrix(0, 8) = "Capital Social y Reservas"
'    Me.Flex.TextMatrix(0, 9) = "Porcentaje Grupo familiar"
'    Me.Flex.TextMatrix(0, 10) = "Limite Maximo Grupo Familiar"
'    Me.Flex.TextMatrix(0, 11) = "Exceden Limite"
'    Me.Flex.TextMatrix(0, 12) = "Monto Exedente"
'
    '*** PEAC 20090402
    Flex.Cols = 17
    Flex.Rows = 1
    Me.Flex.TextMatrix(0, 0) = "Cod Empleado"
    Me.Flex.TextMatrix(0, 1) = "Nombre"
    Me.Flex.TextMatrix(0, 2) = "CodigoPersona"
    Me.Flex.TextMatrix(0, 3) = "Relac. Institucion"
    Me.Flex.TextMatrix(0, 4) = "Nombre Fam"
    Me.Flex.TextMatrix(0, 5) = "Relacion Familiar"
    Me.Flex.TextMatrix(0, 6) = "Monto Creditos Otorgados"
    Me.Flex.TextMatrix(0, 7) = "Saldos de Creditos"
    Me.Flex.TextMatrix(0, 8) = "Monto Consolidado por Grupo Familiar"
    Me.Flex.TextMatrix(0, 9) = "Capital Social y Reservas"
    Me.Flex.TextMatrix(0, 10) = "Porcentaje Grupo familiar"
    Me.Flex.TextMatrix(0, 11) = "Limite Maximo Grupo Familiar"
    Me.Flex.TextMatrix(0, 12) = "Excedente Limite" 'WIOR 20140123 Exceden A Excedente
    Me.Flex.TextMatrix(0, 13) = "Supera Límite" 'WIOR 20140123
    Me.Flex.TextMatrix(0, 14) = "Monto Exedente" 'WIOR 20140123 13 a 14
    'ALPA 20101002**************************************************
    'Me.Flex.TextMatrix(0, 14) = ""
    Me.Flex.TextMatrix(0, 15) = "Fecha Desembolso" 'WIOR 20140123 14 a 15
    Me.Flex.TextMatrix(0, 16) = "Monto Desembolso" 'WIOR 20140123 15 a 16
    '***************************************************************
    While Not rs.EOF
'        Flex.Rows = Flex.Rows + 1
'        Me.Flex.TextMatrix(Flex.Rows - 1, 0) = rs!cRHCod
'        Me.Flex.TextMatrix(Flex.Rows - 1, 1) = rs!NEmp
'        Me.Flex.TextMatrix(Flex.Rows - 1, 2) = rs!cPersCod
'        Me.Flex.TextMatrix(Flex.Rows - 1, 3) = rs!NFam
'        Me.Flex.TextMatrix(Flex.Rows - 1, 4) = rs!cNomTab
'        Me.Flex.TextMatrix(Flex.Rows - 1, 5) = "0.00"
'        Me.Flex.TextMatrix(Flex.Rows - 1, 6) = "0.00"
        
        '*** PEAC 20090402
        Flex.Rows = Flex.Rows + 1
        Me.Flex.TextMatrix(Flex.Rows - 1, 0) = rs!cRHCod
        Me.Flex.TextMatrix(Flex.Rows - 1, 1) = rs!NEmp
        Me.Flex.TextMatrix(Flex.Rows - 1, 2) = rs!cPersCod
        Me.Flex.TextMatrix(Flex.Rows - 1, 3) = rs!cRelaInsti
        Me.Flex.TextMatrix(Flex.Rows - 1, 4) = rs!NFam
        Me.Flex.TextMatrix(Flex.Rows - 1, 5) = rs!cNomtab
        Me.Flex.TextMatrix(Flex.Rows - 1, 6) = "0.00"
        Me.Flex.TextMatrix(Flex.Rows - 1, 7) = "0.00"
        
        If Flex.Rows = 2 Then
            lsCadPers = "'" & rs!cPersCod & "'"
        Else
            lsCadPers = lsCadPers & ",'" & rs!cPersCod & "'"
        End If
        
        rs.MoveNext
    Wend

    rs.Close

    For lnI = 0 To Me.lstAge.ListCount - 1
        If lstAge.Selected(lnI) Then
            Caption = lstAge.List(lnI)
            
            If Me.chkUltFinMes.value = 0 Then
                 Set rs = clsRep.GetCreditosVinculados(lbSoloTitulares, False, lsCadPers, gnTipCambio, Right(lstAge.List(lnI), 2))
                
                While Not rs.EOF
                    For lnJ = 1 To Me.Flex.Rows - 1
                        
'                        If Me.Flex.TextMatrix(lnJ, 2) = rs.Fields(0) Then
'                            Me.Flex.TextMatrix(lnJ, 5) = Format(CCur(Me.Flex.TextMatrix(lnJ, 5)) + rs!Desem, "#0.00")
'                            Me.Flex.TextMatrix(lnJ, 6) = Format(CCur(Me.Flex.TextMatrix(lnJ, 6)) + rs!Saldo, "#0.00")
'                            lnJ = Me.Flex.Rows - 1
'                        End If
                        
                        '*** PEAC 20090402
                        If Me.Flex.TextMatrix(lnJ, 2) = rs.Fields(0) Then
                            Me.Flex.TextMatrix(lnJ, 6) = Format(CCur(Me.Flex.TextMatrix(lnJ, 6)) + rs!Desem, "#0.00")
                            Me.Flex.TextMatrix(lnJ, 7) = Format(CCur(Me.Flex.TextMatrix(lnJ, 7)) + rs!Saldo, "#0.00")
                            Me.Flex.TextMatrix(lnJ, 15) = Format(rs!dVigencia, "YYYY/MM/DD") 'WIOR 20140123 14 a 15
                            Me.Flex.TextMatrix(lnJ, 16) = Format(rs!nMontoEntre, "#0.00") 'WIOR 20140123 15 a 16
                            
                            lnJ = Me.Flex.Rows - 1
                        End If
                    
                    Next lnJ
                    
                    rs.MoveNext
                Wend
                rs.Close
            Else
                Set rs = clsRep.GetCreditosVinculados(lbSoloTitulares, True, lsCadPers, gnTipCambio, Right(lstAge.List(lnI), 2))
            
                While Not rs.EOF
                    For lnJ = 1 To Me.Flex.Rows - 1
                    
'                        If Me.Flex.TextMatrix(lnJ, 2) = rs.Fields(0) Then
'                            Me.Flex.TextMatrix(lnJ, 5) = Format(CCur(Me.Flex.TextMatrix(lnJ, 5)) + rs!Desem, "#0.00")
'                            Me.Flex.TextMatrix(lnJ, 6) = Format(CCur(Me.Flex.TextMatrix(lnJ, 6)) + rs!Saldo, "#0.00")
'                            lnJ = Me.Flex.Rows - 1
'                        End If

                        '*** PEAC 20090402
                        If Me.Flex.TextMatrix(lnJ, 2) = rs.Fields(0) Then
                            Me.Flex.TextMatrix(lnJ, 6) = Format(CCur(Me.Flex.TextMatrix(lnJ, 6)) + rs!Desem, "#0.00")
                            Me.Flex.TextMatrix(lnJ, 7) = Format(CCur(Me.Flex.TextMatrix(lnJ, 7)) + rs!Saldo, "#0.00")
                            Me.Flex.TextMatrix(lnJ, 15) = Format(rs!dVigencia, "YYYY/MM/DD") 'WIOR 20140123 14 a 15
                            Me.Flex.TextMatrix(lnJ, 16) = Format(rs!nMontoEntre, "#0.00") 'WIOR 20140123 15 a 16
                            lnJ = Me.Flex.Rows - 1
                        End If

                    Next lnJ
                    
                    rs.MoveNext
                Wend
                
                rs.Close
            End If
        End If
        Me.prgBar.value = Me.prgBar.value + 1
        DoEvents
    Next lnI
    
    lnI = 1
    lsCodAnt = ""
    Me.Flex.Rows = Me.Flex.Rows + 1
    While Me.Flex.TextMatrix(lnI, 0) <> ""
        If lnI = 1 Then
            lnPosAnt = lnI
            
            '*** PEAC 20090402
            'lnAcumulador = CCur(Me.Flex.TextMatrix(lnI, 6))
            lnAcumulador = CCur(Me.Flex.TextMatrix(lnI, 7))
        End If
                
        If Me.Flex.TextMatrix(lnI, 0) <> lsCodAnt And lnI <> 1 Then
            
'            Me.Flex.TextMatrix(lnPosAnt, 7) = Format(lnAcumulador, "#,##0.00")
'            Me.Flex.TextMatrix(lnPosAnt, 8) = Format(Me.txtMonto.Text, "#,##0.00")
'            Me.Flex.TextMatrix(lnPosAnt, 9) = Format((lnAcumulador * 100) / CCur(Me.txtMonto.Text), "#,##0.00")
'            Me.Flex.TextMatrix(lnPosAnt, 10) = Format(Me.txtPorpar.Text, "#,##0.00")
'            Me.Flex.TextMatrix(lnPosAnt, 11) = Format(CCur(Me.txtPorpar.Text) - CCur(Me.Flex.TextMatrix(lnPosAnt, 9)), "#,##0.00")
'            Me.Flex.TextMatrix(lnPosAnt, 12) = Format((CCur(Me.Flex.TextMatrix(lnPosAnt, 11)) * CCur(txtMonto.Text)) / 100, "#,##0.00")
'
            '*** PEAC 20090402
            Me.Flex.TextMatrix(lnPosAnt, 8) = Format(lnAcumulador, "#,##0.00")
            Me.Flex.TextMatrix(lnPosAnt, 9) = Format(Me.txtMonto.Text, "#,##0.00")
            Me.Flex.TextMatrix(lnPosAnt, 10) = Format((lnAcumulador * 100) / CCur(Me.txtMonto.Text), "#,##0.00")
            Me.Flex.TextMatrix(lnPosAnt, 11) = Format(Me.txtPorpar.Text, "#,##0.00")
            Me.Flex.TextMatrix(lnPosAnt, 12) = Format(CCur(Me.txtPorpar.Text) - CCur(Me.Flex.TextMatrix(lnPosAnt, 10)), "#,##0.00")
            Me.Flex.TextMatrix(lnPosAnt, 13) = IIf(IsNumeric(Flex.TextMatrix(lnPosAnt, 12)), IIf(CDbl(Flex.TextMatrix(lnPosAnt, 12)) < 0, "SI", "NO"), "") 'WIOR 20140123
            Me.Flex.TextMatrix(lnPosAnt, 14) = Format((CCur(Me.Flex.TextMatrix(lnPosAnt, 12)) * CCur(txtMonto.Text)) / 100, "#,##0.00") 'WIOR 20140123 13 a 14
            
            lnPosAnt = lnI
            
            '*** PEAC 20090402
            'lnAcumulador = CCur(Me.Flex.TextMatrix(lnI, 6))
            lnAcumulador = CCur(Me.Flex.TextMatrix(lnI, 7))
            
        End If
        
        If Me.Flex.TextMatrix(lnI, 0) = lsCodAnt Then
            
            'lnAcumulador = lnAcumulador + CCur(Me.Flex.TextMatrix(lnI, 6))
            
            '*** PEAC 20090402
            lnAcumulador = lnAcumulador + CCur(Me.Flex.TextMatrix(lnI, 7))
        End If
        
        lsCodAnt = Me.Flex.TextMatrix(lnI, 0)
        lnI = lnI + 1
    Wend
    
'    Me.Flex.TextMatrix(lnPosAnt, 7) = Format(lnAcumulador, "#,##0.00")
'    Me.Flex.TextMatrix(lnPosAnt, 8) = Format(Me.txtMonto.Text, "#,##0.00")
'    Me.Flex.TextMatrix(lnPosAnt, 9) = Format((lnAcumulador * 100) / CCur(Me.txtMonto.Text), "#,##0.00")
'    Me.Flex.TextMatrix(lnPosAnt, 10) = Format(Me.txtPorpar.Text, "#,##0.00")
'    Me.Flex.TextMatrix(lnPosAnt, 11) = Format(CCur(Me.txtPorpar.Text) - CCur(Me.Flex.TextMatrix(lnPosAnt, 9)), "#,##0.00")
'    Me.Flex.TextMatrix(lnPosAnt, 12) = Format((CCur(Me.Flex.TextMatrix(lnPosAnt, 11)) * CCur(txtMonto.Text)) / 100, "#,##0.00")
    
   '*** PEAC 20090402
    Me.Flex.TextMatrix(lnPosAnt, 8) = Format(lnAcumulador, "#,##0.00")
    Me.Flex.TextMatrix(lnPosAnt, 9) = Format(Me.txtMonto.Text, "#,##0.00")
    Me.Flex.TextMatrix(lnPosAnt, 10) = Format((lnAcumulador * 100) / CCur(Me.txtMonto.Text), "#,##0.00")
    Me.Flex.TextMatrix(lnPosAnt, 11) = Format(Me.txtPorpar.Text, "#,##0.00")
    Me.Flex.TextMatrix(lnPosAnt, 12) = Format(CCur(Me.txtPorpar.Text) - CCur(Me.Flex.TextMatrix(lnPosAnt, 10)), "#,##0.00")
    Me.Flex.TextMatrix(lnPosAnt, 13) = IIf(IsNumeric(Flex.TextMatrix(lnPosAnt, 12)), IIf(CDbl(Flex.TextMatrix(lnPosAnt, 12)) < 0, "SI", "NO"), "") 'WIOR 20140123
    Me.Flex.TextMatrix(lnPosAnt, 14) = Format((CCur(Me.Flex.TextMatrix(lnPosAnt, 12)) * CCur(txtMonto.Text)) / 100, "#,##0.00") 'WIOR 20140123 13 a 14
    
    lnI = 1
    
    While lnI <> Flex.Rows - 1
'        If CCur(Me.Flex.TextMatrix(lnI, 6)) = 0 Then 'And Me.Flex.TextMatrix(lnI, 7) = "" Then
        If CCur(Me.Flex.TextMatrix(lnI, 7)) = 0 Then
            Flex.RemoveItem lnI
            lnI = lnI - 1
'        ElseIf CCur(Me.Flex.TextMatrix(lnI, 6)) = 0 And Not IsNumeric(Me.Flex.TextMatrix(lnI, 7)) Then
        ElseIf CCur(Me.Flex.TextMatrix(lnI, 7)) = 0 And Not IsNumeric(Me.Flex.TextMatrix(lnI, 8)) Then
            Flex.RemoveItem lnI
            lnI = lnI - 1
        End If
        lnI = lnI + 1
    Wend
    
    lnI = 1
    lsCodAnt = ""
    While lnI <> Flex.Rows - 1
        If lsCodAnt = Me.Flex.TextMatrix(lnI, 0) Then
            Me.Flex.TextMatrix(lnI, 0) = ""
            Me.Flex.TextMatrix(lnI, 1) = ""
        Else
            lsCodAnt = Me.Flex.TextMatrix(lnI, 0)
        End If
        Me.Flex.TextMatrix(lnI, 2) = "'" & Me.Flex.TextMatrix(lnI, 2)
        
        lnI = lnI + 1
    Wend
    
On Error GoTo ErrHandler
    Me.prgBar.value = Me.prgBar.Max
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
      
       Call GeneraReporte

       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
     
    Caption = lsCaption
    
    Me.prgBar.value = 0
    Exit Sub
ErrHandler:
    MsgBox "Si desea ver el archivo, genere nuevamente el proceso usted debe presionar 'SI'", vbInformation, "AVISO"
    Me.prgBar.value = 0
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'ORCR20140314**************************
    Dim rs As New ADODB.Recordset
    'Set rs = New ADODB.Recordset
    Dim oCons As New DConstante
    'Set oCons = New DConstante
    '***********************************************
    Dim oConsSist As New COMDConstSistema.NCOMConstSistema
    Dim oPar As New COMDCredito.DCOMParametro
    Dim oPatrimonioEfectivo As New COMNCredito.NCOMPatrimonioEfectivo
    
    Dim fFecActual As Date
    Dim n7PorcPatriEfec As Double
    Dim n5PorcDel7PorcPatriEfec As Double
    Dim nPatrimonioEfec As Double
     '***********************************************
    
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Set rs = oCons.getAgencias(, , True)
    
    Me.prgBar.Max = rs.RecordCount + 1
    
    Me.lstAge.Clear
    
    While Not rs.EOF
        lstAge.AddItem rs.Fields(1) & Space(50) & rs.Fields(0)
        rs.MoveNext
    Wend
    
'    Me.txtPorpar.Text = "5.00"
'    Me.txtporTot.Text = "7.00"
'    Caption = lsCaption
    
    'WIOR 20140123 **************
'    Dim sAnio As String
'    Dim sMes As String
'    Dim oConsSist As COMDConstSistema.NCOMConstSistema
'    Set oConsSist = New COMDConstSistema.NCOMConstSistema
'    fgFecActual = oConsSist.LeeConstSistema(gConstSistCierreMesNegocio)
'
'    sAnio = Year(fgFecActual)
'    sMes = Format(Month(fgFecActual), "00")
'    txtMonto.Text = Format(ObtenerSaldos(sAnio, sMes), "###," & String(15, "#") & "#0.00")
'    txtMonto.Enabled = False
    'WIOR FIN *******************
   
    fFecActual = oConsSist.LeeConstSistema(gConstSistFechaInicioDia)
    
    nPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(fFecActual), Format(Month(fFecActual), "00"))
    
    If nPatrimonioEfec = 0 Then
        fFecActual = DateAdd("m", -1, fFecActual)
        nPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(fFecActual), Format(Month(fFecActual), "00"))
        
        If nPatrimonioEfec = 0 Then
            MsgBox "favor de definir el Patrimonio Efectivo para continuar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    

    n7PorcPatriEfec = oPar.RecuperaValorParametro(102752) '/ 100
    n5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) '/ 100
    
    Me.txtMonto.Text = Format(nPatrimonioEfec, "###," & String(15, "#") & "#0.00") & " "
    Me.txtPorpar.Text = Format(n5PorcDel7PorcPatriEfec, "###," & String(15, "#") & "#0.00") & " "  '"5.00"
    Me.txtporTot.Text = Format(n7PorcPatriEfec, "###," & String(15, "#") & "#0.00") & " " '"7.00"
    Caption = lsCaption
    '***********************************************
    '***********************************************
    
    'ORCR20140314**************************
End Sub

Private Sub GeneraReporte()
    Dim i As Integer
    Dim k As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lnAcum As Currency
    
    Dim sTipoGara As String
    Dim sTipoCred As String
   
    For i = 0 To Me.Flex.Rows - 1
        lnAcum = 0
        For J = 0 To Me.Flex.Cols - 1
            If IsNumeric(Me.Flex.TextMatrix(i, J)) Then
                xlHoja1.Cells(i + 1, J + 1) = Format(Me.Flex.TextMatrix(i, J), "#,##0.00")
            Else
                xlHoja1.Cells(i + 1, J + 1) = Me.Flex.TextMatrix(i, J)
            End If
            If i > 1 And J > 1 Then
                If Me.Flex.TextMatrix(i, J) <> "" Then
                    If IsNumeric(Me.Flex.TextMatrix(i, J)) Then
                        lnAcum = lnAcum + CCur(Me.Flex.TextMatrix(i, J))
                    End If
                End If
            End If
        Next J
        If i > 1 Then
            'VSQL = Format(lnAcum, "#,##0.00")  ' "=SUMA(" & Trim(ExcelColumnaString(3)) & Trim(I + 1) & ":" & Trim(ExcelColumnaString(Me.Flex.Cols)) & Trim(I + 1) & ")"
            'xlHoja1.Cells(I + 1, Me.Flex.Cols + 1).Formula = VSQL
            'xlHoja1.Cells(I + 1, Me.Flex.Cols + 1) = VSQL
        End If
    Next i
        
    xlHoja1.Range("A1:A" & Trim(str(Me.Flex.Rows))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(str(Me.Flex.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True
    xlHoja1.Cells.Select
    xlHoja1.Cells.EntireColumn.AutoFit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'CierraConexion
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 16)
End Sub

Private Sub txtPorpar_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPorpar, KeyAscii, 16)
End Sub


Private Sub txtporTot_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtporTot, KeyAscii, 16)
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

'WIOR 20140123 *********************************************
'Private Function ObtenerSaldos(ByVal psAnio As String, ByVal psMes As String) As Double
'    Dim oNContabilidad As COMNContabilidad.NCOMContFunciones
'    Dim nSaldo As Double
'    Set oNContabilidad = New COMNContabilidad.NCOMContFunciones
'
'    nSaldo = oNContabilidad.PatrimonioEfecAjustInfl(psAnio, psMes)
'    ObtenerSaldos = nSaldo
'    Set oNContabilidad = Nothing
'End Function
'WIOR FIN **************************************************

