VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProvisionGratificacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provision de Gratificaciones"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmProvisionGratificacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConsol 
      Caption         =   "&Consol Excel"
      Height          =   345
      Left            =   1920
      TabIndex        =   20
      Top             =   6600
      Width           =   1260
   End
   Begin VB.TextBox txtEsSalud 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   6600
      Width           =   1100
   End
   Begin VB.TextBox txtGrati 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CheckBox chkEsSalud 
      Caption         =   "Es Salud"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   720
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkPeriodo 
      Caption         =   "Ultimo Mes del Periodo (Asiento Contable)"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton CmdCarga 
      Caption         =   "&Carga"
      Height          =   330
      Left            =   8040
      TabIndex        =   12
      Top             =   195
      Width           =   1155
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8070
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   6585
      Width           =   1560
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   8535
      TabIndex        =   4
      Top             =   6945
      Width           =   1155
   End
   Begin VB.CommandButton cmdProvisionar 
      Caption         =   "&Provisionar"
      Height          =   330
      Left            =   7395
      TabIndex        =   1
      Top             =   6945
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton CmdExcel3 
      Caption         =   "<< Exportar Excel >> "
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   6615
      Visible         =   0   'False
      Width           =   1725
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9446
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Nombre-Grati-Fec. Ing.-Total Ingreso-Gratificacion-Es Salud-Validar-Dias Diff"
      EncabezadosAnchos=   "500-800-4000-800-1200-1200-1300-900-800-500"
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
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-3-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-4-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-R-R-R-R-R-L"
      FormatosEdit    =   "0-0-0-0-0-2-2-2-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   7005
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   315
      Left            =   6900
      TabIndex        =   7
      Top             =   210
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdCorta 
      Enabled         =   0   'False
      Height          =   330
      Left            =   6255
      TabIndex        =   6
      Top             =   6945
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento"
      Height          =   330
      Left            =   5115
      TabIndex        =   9
      Top             =   6945
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Es Salud :"
      Height          =   240
      Left            =   5280
      TabIndex        =   19
      Top             =   6600
      Width           =   705
   End
   Begin VB.Label Label3 
      Caption         =   "Gratif. :"
      Height          =   240
      Left            =   3360
      TabIndex        =   18
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CAJA MAYNAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   6000
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lnlTot 
      Caption         =   "Total :"
      Height          =   240
      Left            =   7560
      TabIndex        =   11
      Top             =   6600
      Width           =   465
   End
   Begin VB.Label lblProceso 
      Caption         =   "Proceso :"
      Height          =   240
      Left            =   6135
      TabIndex        =   8
      Top             =   255
      Width           =   810
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "PROVISION DE GRATIFICACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9555
   End
End
Attribute VB_Name = "frmProvisionGratificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''Private Sub cmdAsiento_Click()
'''    Dim sql As String
'''    Dim Rs As New ADODB.Recordset
'''    Dim ldFechaAnt As Date
'''    Dim rsMontoAnt As New ADODB.Recordset
'''    Dim oCon As New DConecta
'''    Dim oMov As New DMov
'''    Dim lnMovNro As Long
'''    Dim lsMovNro As String
'''    Dim lnItem As Long
'''    Dim lnDiff As Currency
'''
'''    Dim oAsi As NContImprimir
'''    Set oAsi = New NContImprimir
'''    Dim lsCadena As String
'''
'''    Dim lnAcum As Currency
'''    Dim lnAcumDiff As Currency
'''    Dim lnUltimo  As Currency
'''
'''    Dim oPrevio As Previo.clsPrevio
'''    Set oPrevio = New Previo.clsPrevio
'''
'''    ldFechaAnt = DateAdd("d", -1, CDate("01/" & Format(CDate(Me.mskFecha), "mm/yyyy")))
'''    sql = "Select cMovNro  from mov where cmovnro like '" & Format(ldFechaAnt, gsFormatoMovFecha) & "%' And cOpeCod = '622101' And nMovflag = 0"
'''    oCon.AbreConexion
'''
'''    Set Rs = oCon.CargaRecordSet(sql)
'''
'''    If Not Rs.EOF And Not Rs.BOF Then
'''        lsMovNro = Rs!cMovNro
'''        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
'''
'''        lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
'''        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
'''
'''        oCon.CierraConexion
'''        Exit Sub
'''    Else
'''        Rs.Close
'''    End If
'''
'''    If MsgBox("Desea Generar Asiento Contable ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'''
'''    sql = "Select dbo.getsaldocta('" & Format(ldFechaAnt, gsFormatoFecha) & "','21160903',1)"
'''    Set rsMontoAnt = oCon.CargaRecordSet(sql)
'''
''''    Sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionGratificacionesMes('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "') a" _
''''        & " full outer join dbo.RHGetProvisionGratificacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "'," & rsMontoAnt.Fields(0) & ") b on a.cage = b.cage"
''''    Set Rs = oCon.CargaRecordSet(Sql)
'''
'''    If Me.chkPeriodo.value = 1 Then
'''        sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionGratificacionesMes('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "',1) a" _
'''            & " full outer join dbo.RHGetProvisionGratificacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "'," & rsMontoAnt.Fields(0) & ") b on a.cage = b.cage"
'''        Set Rs = oCon.CargaRecordSet(sql)
'''    Else
'''        sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionGratificacionesMes('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "',0) a" _
'''            & " full outer join dbo.RHGetProvisionGratificacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "'," & rsMontoAnt.Fields(0) & ") b on a.cage = b.cage"
'''        Set Rs = oCon.CargaRecordSet(sql)
'''    End If
'''
'''    lsMovNro = oMov.GeneraMovNro(ldFechaAnt, Right(gsCodAge, 2), gsCodUser)
'''    lnAcum = 0
'''    lnItem = 0
'''    lnAcumDiff = 0
'''
'''    oMov.BeginTrans
'''        oMov.InsertaMov lsMovNro, "622101", "Provison de Planilla de Gratificaciones. " & Format(ldFechaAnt, gsFormatoFechaView)
'''        lnMovNro = oMov.GetnMovNro(lsMovNro)
'''
'''        While Not Rs.EOF
'''            lnAcum = lnAcum + Round(Rs!Diff, 2) + Round(Rs!MontoAnt, 2)
'''            lnDiff = Round(Rs!MontoAct - ((Round(Rs!MontoAct, 2) / CCur(Me.txtTotal.Text)) * rsMontoAnt.Fields(0)), 2)
'''            lnAcumDiff = lnAcumDiff + lnDiff
'''            lnItem = lnItem + 1
'''            oMov.InsertaMovCta lnMovNro, lnItem, "45110103" & Rs!Age, lnDiff
'''            lnUltimo = Rs!Diff
'''            Rs.MoveNext
'''        Wend
'''
'''        'If lnAcum <> CCur(Me.txtTotal.Text) Then
'''        '    oMov.ActualizaMovCta lnItem, lnUltimo + (CCur(Me.txtTotal.Text) - lnAcum)
'''        '    lnAcumDiff = lnAcumDiff + (CCur(Me.txtTotal.Text) - lnAcum)
'''        'End If
'''
'''        lnItem = lnItem + 1
'''        oMov.InsertaMovCta lnMovNro, lnItem, "21160903", lnAcumDiff * -1
'''
'''    oMov.CommitTrans
'''
'''    lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
'''
'''    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
'''    oCon.CierraConexion
'''
'''End Sub

Private Sub cmdAsiento_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim ldFechaAnt As Date
    Dim ldFechaMes As Date
    Dim rsMontoAnt As New ADODB.Recordset
    Dim oCon As New DConecta
    Dim oMov As New DMov
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnItem As Long
    
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    Dim lsCadena As String
    
    Dim lnAcum As Currency
    Dim lnAcumDiff As Currency
    Dim lnDiffEsSaludAport As Currency
    Dim lnUltimo  As Currency
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    ldFechaAnt = DateAdd("d", -1, CDate("01/" & Format(CDate(Me.mskFecha), "mm/yyyy")))
    ldFechaMes = Format(CDate(Me.mskFecha), "dd/mm/yyyy")
    
    sql = "Select cMovNro  from mov where cmovnro like '" & Format(ldFechaMes, gsFormatoMovFecha) & "%' And cOpeCod = '622101' And nMovflag = 0"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
     
    If Not rs.EOF And Not rs.BOF Then
        lsMovNro = rs!cMovNro
        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
        
        lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
        
        oCon.CierraConexion
        Exit Sub
    Else
        rs.Close
    End If
    
    If MsgBox("Desea Generar Asiento Contable ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
'    sql = "Select dbo.GetsaldoctaAcumulado('" & Format(ldFechaAnt, gsFormatoFecha) & "','25150402%',1)"
'    sql = "Select 0"
'    Set rsMontoAnt = oCon.CargaRecordSet(sql)
'
'    sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionGratificacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), "yyyymm") & "', " & CCur(Me.txtTotal.Text) & ") a" _
'        & " full outer join dbo.RHGetProvisionGratificacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), "yyyymm") & "'," & rsMontoAnt.Fields(0) & ") b on a.cage = b.cage"
    
        
    sql = "Select cAgenciaActual Age, Sum(Round(ISNULL(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 4)+ SubString(Convert(varchar (10), dIngreso, 112), 5, 2) = '" & Mid(Format(CDate(mskFecha), gsFormatoMovFecha), 1, 6) & "'"
    sql = sql & " Then Round(Round(nRHSueldoMonto / 6, 2) * Round(Cast(Datediff(d, dIngreso, '" & Format(CDate(mskFecha), gcFormatoFecha) & "') as Float )/30, 2, 1), 2)"
    sql = sql & " Else Convert(Decimal (20, 2),(nRHSueldoMonto / 6))End , 0), 2)) nRHGratiMonto,"
    sql = sql & " Sum(Cast((Round(ISNULL(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 4)+ SubString(Convert(varchar (10), dIngreso, 112), 5, 2) = '" & Mid(Format(CDate(mskFecha), gsFormatoMovFecha), 1, 6) & "'"
    sql = sql & " Then Round(Round(nRHSueldoMonto / 6, 2) * Round(Cast(Datediff(d, dIngreso, '" & Format(CDate(mskFecha), gcFormatoFecha) & "') as Decimal (20, 2) )/30, 2, 1), 2)"
    sql = sql & " Else Convert(Decimal (20, 2),(nRHSueldoMonto / 6))End , 0), 2) * 0.09)As Decimal (20, 2))) nRHEsSaludMonto"
    sql = sql & " From("
    sql = sql & " Select PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, Round(Isnull(nRHMesGratificacion, 0),4) nRHMesGratificacion,"
    sql = sql & " IsNull((Select sum(IsNull(nMonto,0)) Monto from RHPlanillaDetCon Where cPersCod = PE.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '" & Mid(Format(CDate(mskFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130') ,0) nRHSueldoMonto, cAgenciaActual "
    sql = sql & " From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod Inner Join RRHH EM On EM.cPersCod = E.cPersCod Inner Join RHConcepto RHPD ON EM.CPERSCOD = RHPD.CPERSCOD "
    sql = sql & " And E.cRRHHPeriodo LIKE '" & Mid(Format(CDate(mskFecha), gsFormatoMovFecha), 1, 6) & "%' And E.cPlanillaCod in ('E01','E08')"
    sql = sql & " Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod GROUP BY PE.CPERSCOD, EM.CRHCOD, PE.CPERSNOMBRE, EM.DINGRESO, cAgenciaActual, nRHMesGratificacion"
    sql = sql & " ) AAA"
    sql = sql & " Group by cAgenciaActual Order by cAgenciaActual"
    
    Set rs = oCon.CargaRecordSet(sql)
    
    lsMovNro = oMov.GeneraMovNro(ldFechaMes, Right(gsCodAge, 2), gsCodUser)
    lnAcum = 0
    lnItem = 0
    lnAcumDiff = 0
    
    oMov.BeginTrans
        oMov.InsertaMov lsMovNro, "622101", "Provisión de Planilla de Gratificaciones. " & Format(ldFechaAnt, gsFormatoFechaView)
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        While Not rs.EOF
            'lnAcum = lnAcum + Round(rs!Diff, 2) + Round(rs!MontoAnt, 2)
            'lnAcumDiff = lnAcumDiff + Round(rs!Diff, 2)
            'lnDiffEsSaludAport = Round(rs!Diff * 0.09, 2)
            
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, "45110103" & rs!Age, rs!nRHGratiMonto 'Round(rs!Diff, 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, "451104" & rs!Age, rs!nRHEsSaludMonto 'lnDiffEsSaludAport
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, "25150402" & rs!Age, (rs!nRHGratiMonto + rs!nRHEsSaludMonto) * -1
            lnUltimo = rs!nRHGratiMonto
            rs.MoveNext
        Wend
        
        'Comentado Por MAVM 20110713 ***
'        If lnAcum <> CCur(Me.txtTotal.Text) Then
'            oMov.ActualizaMovCta lnMovNro, lnItem - 2, , lnUltimo + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")
'            oMov.ActualizaMovCta lnMovNro, lnItem, , (lnUltimo + lnDiffEsSaludAport + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")) * -1
'            lnAcumDiff = lnAcumDiff + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")
'        End If
        '***
        
        'lnItem = lnItem + 1
        'oMov.InsertaMovCta lnMovNro, lnItem, "21160903", lnAcumDiff * -1
    oMov.CommitTrans
        
    lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
    
    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
    oCon.CierraConexion

End Sub


Private Sub CmdCarga_Click()
If ValFecha(Me.mskFecha) = False Then
    Exit Sub
End If
LimpiaFlex
CargaProvisionGrati
'If Month(Me.mskFecha) = 12 Then
'        MsgBox "Se debe provisionar Nomviembre y Diciembre", vbInformation, "AVISO"
'End If
If Month(Me.mskFecha) = 7 Or Month(Me.mskFecha) = 1 Then
    CmdCorta.Caption = Format(Me.mskFecha, "MMMM")
    MsgBox "Importante!" & vbCrLf & "No olvidar de poner en 0 la provision del periodo siguiente", vbInformation, "AVISO"
    Me.CmdCorta.Visible = True
    Me.CmdCorta.Enabled = True
End If
Me.cmdProvisionar.Visible = True
Me.cmdAsiento.Visible = True
'Me.CmdCorta.Visible = True
Me.CmdExcel3.Visible = True
End Sub

Private Sub cmdConsol_Click()
Dim RHG As DRHGratificacion
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i, Y As Integer
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim dProv05Vert, dProv06Vert, dProv07Vert, dProv08Vert, dProv09Vert, dProv10Vert, AA As Double
Dim dProv11Vert, dProv12Vert, dProv01Vert, dProv02Vert, dProv03Vert, dProv04Vert, BB As Double
Dim dProvHorizSum, dProvHorizSumTot, CC As Double

Set RHG = New DRHGratificacion
Set rs = RHG.CargarProvGratiConsol(Mid(Format(mskFecha.Text, gsFormatoMovFecha), 1, 6))

Screen.MousePointer = 11
lsArchivo = "ProvGratiConsol" & Format(Now, "yyyymm") & "_" & Format(Time(), "HHMMSS") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\Spooler\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\Spooler\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("B1") = "CAJA MAYNAS"
xlHoja1.Range("B1:C1").MergeCells = True
xlHoja1.Range("B1").Font.Bold = True
xlHoja1.Range("S1") = gdFecSis
xlHoja1.Range("S2") = gsCodUser
xlHoja1.Range("S2").HorizontalAlignment = xlRight
xlHoja1.Range("H1:H2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Enero"
xlHoja1.Range("E5") = "Febrero"
xlHoja1.Range("F5") = "Marzo"
xlHoja1.Range("G5") = "Abril"
xlHoja1.Range("H5") = "Mayo"
xlHoja1.Range("I5") = "Junio"
xlHoja1.Range("J5") = "Total Ene Jun"
xlHoja1.Range("K5") = "Pago"
xlHoja1.Range("L5") = "Julio"
xlHoja1.Range("M5") = "Agosto"
xlHoja1.Range("N5") = "Setiembre"
xlHoja1.Range("O5") = "Octubre"
xlHoja1.Range("P5") = "Noviembre"
xlHoja1.Range("Q5") = "Diciembre"
xlHoja1.Range("R5") = "Total Jul Dic"
xlHoja1.Range("S5") = "Pago"

xlHoja1.Range("B4:S4").MergeCells = True
xlHoja1.Range("B4") = "PROVISION DE GRATIFICACION " & " " & UCase(Format(mskFecha.Text, "MMMM")) & " DEL " & Format(DateAdd("m", -1, gdFecSis), "YYYY")
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter

xlHoja1.Range("B5:S5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:S5").Interior.ColorIndex = 35
xlHoja1.Range("B5:S5").Font.Bold = True

xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 40
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 12
xlHoja1.Range("G1").ColumnWidth = 12
xlHoja1.Range("H1").ColumnWidth = 12
xlHoja1.Range("I1").ColumnWidth = 12
xlHoja1.Range("J1").ColumnWidth = 17
xlHoja1.Range("K1").ColumnWidth = 17
xlHoja1.Range("L1").ColumnWidth = 12
xlHoja1.Range("M1").ColumnWidth = 12
xlHoja1.Range("N1").ColumnWidth = 12
xlHoja1.Range("O1").ColumnWidth = 12
xlHoja1.Range("P1").ColumnWidth = 12
xlHoja1.Range("Q1").ColumnWidth = 12
xlHoja1.Range("R1").ColumnWidth = 17
xlHoja1.Range("S1").ColumnWidth = 17

xlHoja1.Application.ActiveWindow.Zoom = 80
Y = 6

For i = 1 To rs.RecordCount
   
    xlHoja1.Range("B" & Y) = rs!cRHCod
    xlHoja1.Range("C" & Y) = rs!cPersNombre
    
    If (Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "01" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "02" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "03" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "04" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "05" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "06") Then
        
        xlHoja1.Range("D" & Y) = rs!Prov01
        dProv01Vert = dProv01Vert + rs!Prov01
        
        xlHoja1.Range("E" & Y) = rs!Prov02
        dProv02Vert = dProv02Vert + rs!Prov02
        
        xlHoja1.Range("F" & Y) = rs!Prov03
        dProv03Vert = dProv03Vert + rs!Prov03
        
        xlHoja1.Range("G" & Y) = rs!Prov04
        dProv04Vert = dProv04Vert + rs!Prov04
        
        xlHoja1.Range("H" & Y) = rs!Prov05
        dProv05Vert = dProv05Vert + rs!Prov05
        dProvHorizSum = dProvHorizSum + rs!Prov05
        
        xlHoja1.Range("I" & Y) = rs!Prov06
        dProv06Vert = dProv06Vert + rs!Prov06
        dProvHorizSum = dProvHorizSum + rs!Prov06
    End If
    
    If (Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "07" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "08" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "09" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "10" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "11" Or Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) = "12") Then
                
        xlHoja1.Range("L" & Y) = rs!Prov07
        dProv07Vert = dProv07Vert + rs!Prov07
        dProvHorizSum = dProvHorizSum + rs!Prov07
        
        xlHoja1.Range("M" & Y) = rs!Prov08
        dProv08Vert = dProv08Vert + rs!Prov08
        dProvHorizSum = dProvHorizSum + rs!Prov08
        
        xlHoja1.Range("N" & Y) = rs!Prov09
        dProv09Vert = dProv09Vert + rs!Prov09
        dProvHorizSum = dProvHorizSum + rs!Prov09
        
        xlHoja1.Range("O" & Y) = rs!Prov10
        dProv10Vert = dProv10Vert + rs!Prov10
        dProvHorizSum = dProvHorizSum + rs!Prov10
        
        
        xlHoja1.Range("P" & Y) = rs!Prov11
        dProv11Vert = dProv11Vert + rs!Prov11
        dProvHorizSum = dProvHorizSum + rs!Prov11
        
        xlHoja1.Range("Q" & Y) = rs!Prov12
        dProv12Vert = dProv12Vert + rs!Prov12
        dProvHorizSum = dProvHorizSum + rs!Prov12
        xlHoja1.Range("R" & Y) = Format(dProvHorizSum, "#,##0.00")
        
    End If
  
    dProvHorizSumTot = dProvHorizSumTot + dProvHorizSum
    dProvHorizSum = 0
    
    rs.MoveNext
    Y = Y + 1
Next i

xlHoja1.Cells(rs.RecordCount + 6, 2) = "TOTALES"

xlHoja1.Cells(rs.RecordCount + 6, 12) = Format(dProv07Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 13) = Format(dProv08Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 14) = Format(dProv09Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 15) = Format(dProv10Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 16) = Format(dProv11Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 17) = Format(dProv12Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 6, 4) = Format(dProv01Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 5) = Format(dProv02Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 6) = Format(dProv03Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 7) = Format(dProv04Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 8) = Format(dProv05Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 9) = Format(dProv06Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 6, 18) = Format(dProvHorizSumTot, "#,##0.00")

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente en la carpeta Spooler de SICMACT ADM", vbInformation, "Aviso"

CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
End Sub

Private Sub CmdCorta_Click()
Dim i As Integer
Dim mes As Integer
Dim Fecha As String
Dim FechaReg As String
Dim RHG As DRHGratificacion
Set RHG = New DRHGratificacion

mes = Month(Me.mskFecha)
'If Mes = 7 Then
    Fecha = Format(mskFecha, "YYYYMM")
'Else
'    Fecha = Format(gdFecSis, "YYYYMM")
'End If

For i = 1 To Flex.Rows - 1
    If Flex.TextMatrix(i, 3) = "." Then
        If Not RHG.VerificaMigracionGrati(Flex.TextMatrix(i, 1), Fecha, mes) Then
            If RHG.MigraPeriodoGrati(Flex.TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Fecha, mes, Flex.TextMatrix(i, 6)) Then
                Flex.TextMatrix(i, 8) = "OK"
            Else
                Flex.TextMatrix(i, 8) = "Error"
            End If
        Else
            Flex.TextMatrix(i, 8) = "Ya se migro"
        End If
    Else
        Flex.TextMatrix(i, 8) = "No se selec."
    End If
Next i
CmdCarga_Click
End Sub

Private Sub CmdExcel3_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i As Integer
Dim Y As Integer
Dim sSuma As String
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim sCabecera As String

sCabecera = "ProvGrat"

If Me.Flex.TextMatrix(1, 1) = "" Then
    Exit Sub
End If


'On Error GoTo ErrorINFO4Excel
Screen.MousePointer = 11
lsArchivo = "Prov_Grati_" & Format(Me.mskFecha, "YYYYMM") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = sCabecera & Format(Me.mskFecha, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("A1") = "CAJA MAYNAS"
xlHoja1.Range("A1").Font.Bold = True
xlHoja1.Range("F1") = gdFecSis
xlHoja1.Range("F2") = gsCodUser
xlHoja1.Range("F2").HorizontalAlignment = xlRight
xlHoja1.Range("F1:F2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Fec Ing"
xlHoja1.Range("E5") = "Total Ing."
xlHoja1.Range("F5") = "Mes Comp"
xlHoja1.Range("G5") = "Es Salud"
xlHoja1.Range("H5") = "Total"


xlHoja1.Range("B4:H4").MergeCells = True
xlHoja1.Range("B4") = "PROVISION DE GRATIFICACION DEL MES DE " & UCase(Format(DateAdd("M", -1, gdFecSis), "MMMM"))
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:H5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:H5").Interior.ColorIndex = 35
xlHoja1.Range("B5:H5").Font.Bold = True
xlHoja1.Range("A1").ColumnWidth = 6
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 11
xlHoja1.Range("G1").ColumnWidth = 13

xlHoja1.Application.ActiveWindow.Zoom = 80
'xlHoja1.Range("D6:D1000").NumberFormat = "dd/mm/yyyy"
xlHoja1.Range("E6:E1000").Style = "Comma"
xlHoja1.Range("F6:F1000").Style = "Comma"
xlHoja1.Range("G6:G1000").Style = "Comma"
Y = 6
Me.PrgBar.Min = 1
Me.PrgBar.Max = Flex.Rows - 1

For i = 1 To Flex.Rows - 1
   xlHoja1.Range("A" & Y) = Flex.TextMatrix(i, 0)
   xlHoja1.Range("B" & Y) = Flex.TextMatrix(i, 1)
   xlHoja1.Range("C" & Y) = Flex.TextMatrix(i, 2)
   xlHoja1.Range("D" & Y) = "'" & Flex.TextMatrix(i, 4)
   xlHoja1.Range("E" & Y) = Flex.TextMatrix(i, 5)
   xlHoja1.Range("F" & Y) = Format(Flex.TextMatrix(i, 6), "#,##0.00")
   xlHoja1.Range("G" & Y) = Format(Flex.TextMatrix(i, 7), "#,##0.00")
   xlHoja1.Range("H" & Y) = Format(CDbl(Flex.TextMatrix(i, 6)) + CDbl(Flex.TextMatrix(i, 7)), "#,##0.00")
   Y = Y + 1
   Me.PrgBar.value = i
Next i
    
'xlHoja1.SaveAs App.path & "\SPOOLER\" & "VacPRov"
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente", vbInformation, "Aviso"
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
'ErrorINFO4Excel:
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'    xlLibro.Close
'    ' Cierra Microsoft Excel con el método Quit.
'    xlAplicacion.Quit
'    'Libera los objetos.
'    Set xlAplicacion = Nothing
'    Set xlLibro = Nothing
'    Set xlHoja1 = Nothing
End Sub

Private Sub cmdProvisionar_Click()
Dim i As Integer
Dim RHG As DRHGratificacion
Dim Fecha As String
Dim opt As Integer
Dim nDias As Double
Set RHG = New DRHGratificacion
'Fecha = Format(DateAdd("m", -1, gdFecSis), "YYYYMM")
Fecha = Format(Me.mskFecha, "YYYYMM")
opt = MsgBox("Esta seguro de hacer la Provision ", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub

PrgBar.Min = 1
PrgBar.Max = Flex.Rows - 1
PrgBar.Visible = True
For i = 1 To Flex.Rows - 1
    'If Flex.TextMatrix(i, 3) = "." Then
        If Not RHG.VerificaGratificacionMes(Flex.TextMatrix(i, 1), Fecha) = True Then
'            nDias = Flex.TextMatrix(i, 9)
'            If nDias <> 1 Then
'                MsgBox "x"
'            End If
            If RHG.GrabaGratificacion(Flex.TextMatrix(i, 1), gsCodUser, Fecha, FechaHora(gdFecSis), nDias, (CDbl(Flex.TextMatrix(i, 6)) + CDbl(Flex.TextMatrix(i, 7)))) Then
                Flex.TextMatrix(i, 8) = "OK"
                Flex.TextMatrix(i, 6) = Flex.TextMatrix(i, 6) + 1
            Else
                Flex.TextMatrix(i, 8) = "Error"
            End If
        Else
            Flex.TextMatrix(i, 8) = "Ya se abono"
        End If
    'Else
    '    Flex.TextMatrix(i, 8) = "No Selecionado"
    'End If
    PrgBar.value = i
Next i
PrgBar.Visible = False

MsgBox "Provision Realiada", vbInformation, "AVISO"
LimpiaFlex
CargaProvisionGrati

'If Month(gdFecSis) = "12" Then
'    If RHG.VerificaMes(Format(gdFecSis, "YYYY") & "11") Then
'        If Not RHG.VerificaMes(Format(gdFecSis, "YYYY") & "12") Then
'            Opt = MsgBox("Desea hacer la Provision de Diciembre??", vbQuestion + vbYesNo, "AVISO")
'            If Opt = vbNo Then
'                Exit Sub
'            Else
'                CargaProvisionGrati
'            End If
'        End If
'    End If
'End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FechaP As Date
FechaP = "01" & Mid(gdFecSis, 3, 10)
Me.mskFecha = DateAdd("d", -1, FechaP)
End Sub

Sub CargaProvisionGrati()
Dim RHG As DRHGratificacion
Dim rs As ADODB.Recordset
Dim Fecha As String
Dim FechaTemp As String
Dim FechaFeriado As String
Dim ban As Boolean
Dim sql As String
Dim Co  As DConecta
Set RHG = New DRHGratificacion

Dim nDias As Integer
Dim nCalc As Double
Dim lnMonto As Currency
Dim lnEsSalud As Currency

'Fecha = DateAdd("m", -1, gdFecSis)
Fecha = Format(Me.mskFecha, "YYYYMM")

If Month(mskFecha) = 12 Or Month(mskFecha) = 6 Then
    If Me.chkPeriodo.value <> 1 Then
        MsgBox "No se Marco el Check de Ultimo mes", vbInformation, "AVISO"
        Exit Sub
    End If
    If Month(mskFecha) = 12 Then
        Label2 = "PROVISION DE GRATIFICACION DE DICIEMBRE DEL " & Format(mskFecha, "YYYY")
        'Set rs = RHG.Get_Personal_Grati(Format(mskFecha, "YYYY/MM/DD"), "12")
    Else
        Label2 = "PROVISION DE GRATIFICACION DE JUNIO DEL " & Format(mskFecha, "YYYY")
        Fecha = DateAdd("m", 1, Me.mskFecha)
        'Set rs = RHG.Get_Personal_Grati(Format(mskFecha, "YYYY/MM/DD"), "06")
    End If
    'If Not RHG.VerificaMes(Format(gdFecSis, "YYYY") & "11") Then
    '    Label2 = "PROVISION DE " & UCase(Format(Fecha, "MMMM")) & " DEL " & Format(Fecha, "YYYY")
    '    Set rs = RHG.Get_Personal_Grati(Format(gdFecSis, "YYYY/MM/DD"))
    'Else
    'End If
    
    'Solo para cuando es final del Periodo de Gratificaion

    sql = " Select cRHCod, P.cPersCod, cPersNombre, nMonto, dIngreso , nRHMesGratificacion,"
    sql = sql & " ( Select nRHSueldoMonto from RHSueldo RS where RS.cPersCod = P.cPersCod and"
    sql = sql & "   dRHSueldofecha = (Select Max(dRHSueldofecha) from RHSueldo where cPersCod = P.cPersCod )"
    sql = sql & " ) Sueldo,"
    sql = sql & " ("
    sql = sql & "   Select nRHContratoTpo from RHContrato RC where RC.cPersCod = P.cPersCod"
    sql = sql & "     and  cRHContratoNro = ("
    sql = sql & "       Select max(cRHContratoNro) from RHContrato where  cPersCod = P.cPersCod)) Contrato"
    sql = sql & " from dbo.RHPlanillaDetCon  RP"
    sql = sql & " Inner Join Persona P on P.cPersCod = RP.cPersCod"
    sql = sql & " Inner Join RRHH RH on RH.cPersCod = RP.cPersCod"
    sql = sql & " Inner Join RHEmpleado RE on Re.cPersCod =  RP.cPersCod"
    sql = sql & " where cRRHHPeriodo like '" & Format(Fecha, "YYYYMM") & "%' and cPlanillaCod ='E02' and cRHConceptoCod = '130'"
    sql = sql & " Order by cRHCod"
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(sql)
    Co.CierraConexion
    If Not (rs.EOF And rs.BOF) Then
        PrgBar.Visible = True
        PrgBar.Min = 0
        PrgBar.Max = rs.RecordCount
        lnMonto = 0
        While Not rs.EOF
            FechaTemp = Format(rs!dIngreso, "DD/MM/YYYY")
            ban = True
            If ban Then
                Flex.AdicionaFila
                Flex.TextMatrix(Flex.Rows - 1, 1) = rs!cRHCod
                Flex.TextMatrix(Flex.Rows - 1, 2) = rs!cPersNombre
                If Right(Format(rs!dIngreso, "DD/MM/YYYY"), 7) <> Right(Fecha, 7) Then
                    Flex.TextMatrix(Flex.Rows - 1, 3) = "1"
                End If
                Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!dIngreso, "DD/MM/YYYY")
                Flex.TextMatrix(Flex.Rows - 1, 5) = rs!Sueldo
                Flex.TextMatrix(Flex.Rows - 1, 6) = rs!nRHMesGratificacion
                Flex.TextMatrix(Flex.Rows - 1, 7) = rs!nMonto 'Round((rs!nRHSueldoMonto / 6) * rs!nRHMesGratificacion, 2)
                lnMonto = lnMonto + rs!nMonto 'Round((rs!nRHSueldoMonto / 6) * rs!nRHMesGratificacion, 2)
                
                If Flex.TextMatrix(Flex.Rows - 1, 3) = "." Then
                    Flex.TextMatrix(Flex.Rows - 1, 9) = 1
                Else
                    nDias = Day(Format(rs!dIngreso, "DD/MM/YYYY"))
                    nCalc = (30 - nDias) / 30
                    If nDias <= 0 Then
                        Flex.TextMatrix(Flex.Rows - 1, 9) = 0
                    Else
                        Flex.TextMatrix(Flex.Rows - 1, 9) = Round(nCalc, 2)
                    End If
                    
                End If
            
            End If
            PrgBar.value = rs.Bookmark
            rs.MoveNext
        Wend
        PrgBar.Visible = False
        Me.txtTotal.Text = Format(lnMonto, "#,##0.00")
    End If
    MsgBox "Se Calculo la Provision del " & Me.mskFecha, vbInformation, "AVISO"
    Set rs = Nothing
    Set RHG = Nothing
    Exit Sub
Else
    Label2 = "PROVISION DE GRATIFICACION DE " & UCase(Format(mskFecha, "MMMM")) & " DEL " & Format(mskFecha, "YYYY")
    Set rs = RHG.Get_Personal_Grati(Format(mskFecha, "YYYY/MM/DD"))
End If


If Not (rs.EOF And rs.BOF) Then
    PrgBar.Visible = True
    PrgBar.Min = 0
    PrgBar.Max = rs.RecordCount
    lnMonto = 0
    While Not rs.EOF
        FechaTemp = Format(rs!dIngreso, "DD/MM/YYYY")
        ban = True
        If ban Then
            Flex.AdicionaFila
            Flex.TextMatrix(Flex.Rows - 1, 1) = rs!cRHCod
            Flex.TextMatrix(Flex.Rows - 1, 2) = rs!cPersNombre
            If Format(rs!dIngreso, "YYYYMM") <> Fecha Then
                Flex.TextMatrix(Flex.Rows - 1, 3) = "1"
            End If
            Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!dIngreso, "DD/MM/YYYY")
            Flex.TextMatrix(Flex.Rows - 1, 5) = Format(IIf(IsNull(rs!nRHSueldoMonto), 0, rs!nRHSueldoMonto), "#0.00")
            
'            If Flex.TextMatrix(Flex.Rows - 1, 3) = "." Then
'                Flex.TextMatrix(Flex.Rows - 1, 9) = 1
'            Else
'                nDias = Day(Format(rs!dIngreso, "DD/MM/YYYY"))
'                nCalc = (30 - nDias) / 30
'                If nDias <= 0 Then
'                    Flex.TextMatrix(Flex.Rows - 1, 9) = 0
'                Else
'                    Flex.TextMatrix(Flex.Rows - 1, 9) = Int(Format(nCalc, "#0.000") * 10 ^ 2) / 10 ^ 2
'                End If
'
'            End If
            
            'Flex.TextMatrix(Flex.Rows - 1, 7) = Round((rs!nRHSueldoMonto / 6) * rs!nRHMesGratificacion, 2) 'MAVM 20110715
            'If Flex.TextMatrix(Flex.Rows - 1, 3) = "." Then
                Flex.TextMatrix(Flex.Rows - 1, 6) = Format(IIf(IsNull(rs!nRHGratiMonto), 0, rs!nRHGratiMonto), "#0.00") 'MAVM 20110715
                'Flex.TextMatrix(Flex.Rows - 1, 6) = rs!nRHMesGratificacion 'MAVM 20110715
                If chkEsSalud.value = 1 Then
                    Flex.TextMatrix(Flex.Rows - 1, 7) = IIf(IsNull(rs!nRHEsSaludMonto), 0, Format((rs!nRHEsSaludMonto), "#0.00")) 'MAVM 20110715
                Else
                    Flex.TextMatrix(Flex.Rows - 1, 7) = Format("0", "#0.00") 'MAVM 20110715
                End If
            'Else
'                Flex.TextMatrix(Flex.Rows - 1, 6) = Format(IIf(IsNull(rs!nRHSueldoMonto), 0, Round((rs!nRHGratiMonto) * Flex.TextMatrix(Flex.Rows - 1, 9), 2)), "#0.00") 'MAVM 20110715
'                If chkEsSalud.value = 1 Then
'                    Flex.TextMatrix(Flex.Rows - 1, 7) = IIf(IsNull(rs!nRHSueldoMonto), 0, Format(((rs!nRHGratiMonto) * Flex.TextMatrix(Flex.Rows - 1, 9)) * 0.09, "#0.00")) 'MAVM 20110715
'                Else
'                    Flex.TextMatrix(Flex.Rows - 1, 7) = Format("0", "#0.00") 'MAVM 20110715
'                End If
            'End If
            
            'lnMonto = lnMonto + Round((rs!nRHSueldoMonto / 6) * rs!nRHMesGratificacion, 2) 'MAVM 20110715
            lnMonto = lnMonto + Format(Round((Flex.TextMatrix(Flex.Rows - 1, 6)), 2), "#0.00") 'MAVM 20110715
            lnEsSalud = lnEsSalud + Format(Round((Flex.TextMatrix(Flex.Rows - 1, 7)), 2), "#0.00") 'MAVM 20110715
            
        End If
        PrgBar.value = rs.Bookmark
        rs.MoveNext
    Wend
    PrgBar.Visible = False
    'Me.txtTotal.Text = Format(lnMonto, "#,##0.00")
    Me.txtGrati.Text = Format(lnMonto, "#,##0.00")
    Me.txtEsSalud.Text = Format(lnEsSalud, "#,##0.00")
    Me.txtTotal.Text = Format(lnMonto + lnEsSalud, "#,##0.00")
End If

Set rs = Nothing
Set RHG = Nothing
End Sub

Sub LimpiaFlex()
    Flex.Rows = 2
    Flex.Clear
    Flex.FormaCabecera
End Sub
