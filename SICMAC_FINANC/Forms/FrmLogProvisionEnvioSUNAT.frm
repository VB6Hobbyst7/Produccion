VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmLogProvisionEnvioSUNAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Provisiones a Consulta SUNAT"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "FrmLogProvisionEnvioSUNAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Seleccionar Todos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   4230
      Width           =   1845
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Documentos Emitidos"
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
      Height          =   660
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Width           =   10215
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8730
         TabIndex        =   10
         Top             =   180
         Width           =   1245
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   5190
         TabIndex        =   7
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   1605
         TabIndex        =   9
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   4590
         TabIndex        =   8
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   9090
      TabIndex        =   3
      Top             =   4620
      Width           =   1155
   End
   Begin VB.CommandButton cmdEnvio 
      Caption         =   "&Enviar"
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
      Left            =   7920
      TabIndex        =   2
      Top             =   4620
      Width           =   1155
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8955
      TabIndex        =   0
      Top             =   4230
      Width           =   1290
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3375
      Left            =   60
      TabIndex        =   4
      Top             =   750
      Width           =   10215
      _extentx        =   18018
      _extenty        =   5847
      cols0           =   24
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   $"FrmLogProvisionEnvioSUNAT.frx":08CA
      encabezadosanchos=   "400-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-1400-1500-2000-2300-1500-1800-1200-2500-2500-1200"
      font            =   "FrmLogProvisionEnvioSUNAT.frx":09C3
      font            =   "FrmLogProvisionEnvioSUNAT.frx":09EF
      font            =   "FrmLogProvisionEnvioSUNAT.frx":0A1B
      font            =   "FrmLogProvisionEnvioSUNAT.frx":0A47
      font            =   "FrmLogProvisionEnvioSUNAT.frx":0A73
      fontfixed       =   "FrmLogProvisionEnvioSUNAT.frx":0A9F
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-C-C-L-R-R-R-R-L-R-R-R-R"
      formatosedit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0-0-0-2-0-2-2-0-2-2-2-4"
      textarray0      =   "Nro."
      lbeditarflex    =   -1
      lbpuntero       =   -1
      lbordenacol     =   -1
      colwidth0       =   405
      rowheight0      =   360
      forecolorfixed  =   -2147483630
   End
   Begin VB.Label lblSTot 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7980
      TabIndex        =   1
      Top             =   4290
      Width           =   765
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   7860
      Top             =   4230
      Width           =   2385
   End
End
Attribute VB_Name = "FrmLogProvisionEnvioSUNAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodos_Click()
    Dim i As Integer
    
    If Me.chkTodos.value = 1 Then
        For i = 1 To Me.fg.Rows - 1
            Me.fg.TextMatrix(i, 2) = 1
            fg_OnCellCheck i, 2
        Next i
    Else
        For i = 1 To Me.fg.Rows - 1
            Me.fg.TextMatrix(i, 2) = 0
            fg_OnCellCheck i, 2
        Next i
    End If
End Sub

Private Sub cmdEnvio_Click()
    Dim sSQL        As String
    Dim i           As Integer
    Dim sMovEnvio   As String
    Dim oCon        As New DConecta
    Dim lbVerifica As Boolean
    
    lbVerifica = False
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            lbVerifica = True
            Exit For
        End If
    Next i
    
    If lbVerifica = False Then
        MsgBox " Seleccione los docuementos a enviar ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox(" ¿ Esta seguro de enviar los documentos a consulta SUNAT ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    
    Me.cmdEnvio.Enabled = False
    oCon.AbreConexion
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            sMovEnvio = GeneraMovNroActualiza(gdFecSis, gsCodUser, "112", gsCodAge)
            sSQL = " Update MovProvisionAgencia set cMovEnvio='" & sMovEnvio & "' where nmovnro=" & Me.fg.TextMatrix(i, 10) & ""
            oCon.Ejecutar sSQL
            sMovEnvio = ""
        End If
    Next i
    oCon.CierraConexion
    Me.cmdEnvio.Enabled = True
    Me.cmdProcesar = True
End Sub

Private Sub cmdProcesar_Click()
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
Dim sql As String
Dim lsDocs As String
Dim oOpe As New DOperacion
'Dim oDCaja As New DCajaGeneral

Me.fg.Rows = 2
Me.fg.Clear
Me.fg.FormaCabecera

Me.fg.ColWidth(15) = 0
Me.fg.ColWidth(16) = 0
Me.fg.ColWidth(17) = 0
Me.fg.ColWidth(18) = 0
Me.fg.ColWidth(19) = 0
Me.fg.ColWidth(20) = 0
Me.fg.ColWidth(21) = 0
Me.fg.ColWidth(22) = 0

cCtaDetraTemp = Mid(cCtaDetraccionProvision, 1, 2) & Mid(gsOpeCod, 3, 1) & Mid(cCtaDetraccionProvision, 4, Len(cCtaDetraccionProvision) - 2)
 
Set rs = oOpe.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
lsDocs = RSMuestraLista(rs, 1)
Set oOpe = Nothing
RSClose rs


If Mid(gsOpeCod, 3, 1) = "2" Then
    cCtaDetraTempMNE = Mid(cCtaDetraccionProvision, 1, 2) & "1" & Mid(cCtaDetraccionProvision, 4, Len(cCtaDetraccionProvision) - 2)
    lsCtaContDebeBMNE = ""
Else
    cCtaDetraTempMNE = cCtaDetraTemp
    lsCtaContDebeBMNE = lsCtaContDebeB
End If

'Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsCtaContDebeBMNE & "'", lsDocs, txtfechaDel, txtFechaAl, , 2, cCtaDetraTempMNE)

    cCtaDetraTemp = cCtaDetraTemp & IIf(cCtaDetraTempME = "", "", "," & cCtaDetraTempME)
    
        sql = "SELECT distinct md.dDocFecha, doc.cDocAbrev, md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod, "
        sql = sql & "       m.cMovNro, m.nMovNro, mc.cCtaContCod, ISNULL(me.nMovMeImporte + dbo.getMontoPenalidad(m.nMovNro),mc.nMovImporte+dbo.getMontoPenalidad(m.nMovNro)) * -1 as nMovImporte, (mc.nMovImporte +dbo.getMontoPenalidad(m.nMovNro))* -1 as nMovImporteSoles,Provi.cPersIDnro,dbo.getMontoPenalidad(m.nMovNro) nPenalidad  "
        sql = sql & "   ,isnull(ps.nimportecoactivo,0) nImporteCoactivo, case when me.nMovMeImporte is NULL then (mc.nMovImporte*-1-ps.nimportecoactivo) else Round((me.nMovMeImporte*-1)-(ps.nimportecoactivo/ps.ntpocambio),2) end MontoPago, ps.nTpoCambio ,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,''),0) MontoPagadoSUNAT,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,'1'),0) MontoPagadoSUNATS "
        
        
        
        sql = sql & " FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro "
        sql = sql & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
        sql = sql & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro "
        sql = sql & "       LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
        sql = sql & "                   WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
        sql = sql & "                  ) ref ON  ref.nMovNroRef = m.nMovNro "
        sql = sql & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod "
        sql = sql & "             left join Persid ProvI  ON ProvI.cPersCod = mg.cPersCod and cPersIdTpo=2"
        sql = sql & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo "
        sql = sql & "  "

        sql = sql & "  left join movControlPagoSunat ps ON ps.nmovnro = m.nMovNro and ps.bvigente=1 and bvalido=1 "
        sql = sql & " JOIN MovProvisionAgencia Mpa on Mpa.nmovnro=m.nmovnro and Mpa.cAgeCod='" & gsCodAge & "'"
        
        sql = sql & " WHERE Mpa.cMovEnvio is NULL And m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
        sql = sql & " and  ps.cmovnro is null and md.ndoctpo<>40  "
       
        
        '**************************************   D E T R A C C I O N   **************************************
        
            
            sql = sql & " and M.nMovNro not in( select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and md.nestado=3) "
            
            sql = sql & " and (( "
             
            sql = sql & "   M.nMovNro not in"
            sql = sql & " ( Select C.nMovOrigen "
            sql = sql & "   From"
            sql = sql & "     ( Select C1.nMovNro as nMovOrigen, C2.nMovNro AS nMovDestino FROM"
                sql = sql & " ( select distinct mK.nMovNro "
                sql = sql & " from MovCta MK Inner Join Mov MV on MK.nmovnro=Mv.nmovnro "
                sql = sql & " where MK.cCtaContCod Like '" & cCtaDetraTemp & "' and MK.nMovImporte <0 and MV.nmovflag=0 "
                sql = sql & " ) C1 "
            sql = sql & "       Left Join"
                sql = sql & " (Select MR.nMovNro, MR.nMovNroRef "
                sql = sql & " From MovRef MR Inner Join Mov MV1 on MR.nmovnroref = MV1.nmovnro And IsNull(MR.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10  "
                sql = sql & " Inner Join Mov MV2 on MR.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10 and mv2.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "') "
                sql = sql & "  ) C2 "
            sql = sql & "       On C1.nMovNro=C2.nMovNroRef   where  C2.nMovNroRef is not null ) C"
            sql = sql & "   Group By C.nMovOrigen having Count(C.nMovDestino)<>1"
            
            sql = sql & "   Union "
            sql = sql & "   Select Distinct MR1.nMovNroref"
            sql = sql & "   From MovRef MR1 "
                
                sql = sql & " Inner Join Mov MV1 on MR1.nmovnroref=MV1.nmovnro And IsNull(MR1.cAgeCodRef,'') = '' ANd MV1.nMovFlag = 0 ANd MV1.nMovEstado = 10 and mv1.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "','" & OpeCGOpeProvPagoSUNAT & "') " '
                sql = sql & " Inner Join Mov MV2 on MR1.nmovnro = mv2.nMovNro And mv2.nMovFlag = 0 And mv2.nMovEstado = 10  and mv2.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "','" & OpeCGOpeProvPagoSUNAT & "') " '
            
            sql = sql & "   Inner Join MovDoc MD on MR1.nMovNroRef=MD.nMovNro"
            sql = sql & "   where Len(MR1.nMovNro) > 0 "
            
                sql = sql & " and MV1.nMovFlag = 0 And mv2.nMovFlag = 0 "
            
            sql = sql & " and MR1.nMovNroref not in ("
            sql = sql & "   select distinct movcta.nMovNro"
            sql = sql & "   From MovCta  Inner Join Mov On Movcta.nMovNro= Mov.nMovNro "
            sql = sql & "   where cCtaContCod = '29" & Mid(gsOpeCod, 3, 1) & "80799' and nMovImporte < 0 and Mov.nMovFlag=0 "
            sql = sql & "   )"
            
            sql = sql & " )) or (M.nMovNro in(select M.nMovNro From Mov M Inner Join MovDetra MD On M.nMovNro=MD.nMovNro Where M.nMovFlag=0 and MD.nEstado=2))) "
                
        '**************************************   F I N   D E T R A C C I O N   **************************************
        'If psTipoInterfaz = "RECHAZO" Or psTipoInterfaz = "lbReporte" Then
            

        sql = sql & "  and md.dDocFecha BETWEEN '" & Format(Me.txtFechaDel.Text, "mm/dd/yyyy") & "' AND '" & Format(Me.txtFechaAl.Text, "mm/dd/yyyy") & "' "
        sql = sql & "    and md.nDocTpo IN (" & lsDocs & ") and mc.cCtaContCod IN ('25" & Mid(gsOpeCod, 3, 1) & "601','25" & Mid(gsOpeCod, 3, 1) & "602') and m.copecod not in ('" & OpeRegPenalidadMN & "','" & OpeRegPenalidadME & "') "
        sql = sql & " ORDER BY Prov.cpersnombre, m.nMovNro"
        
  oCon.AbreConexion
  Set rs = oCon.CargaRecordSet(sql)
  Do While Not rs.EOF
        fg.AdicionaFila
        nItem = fg.Row
        
        fg.TextMatrix(nItem, 1) = nItem
        fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
        fg.TextMatrix(nItem, 4) = rs!dDocFecha
        fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPersona, True)
        fg.TextMatrix(nItem, 6) = rs!cMovDesc
        fg.TextMatrix(nItem, 7) = Format(rs!nmovimporte, gsFormatoNumeroView)
        fg.TextMatrix(nItem, 8) = rs!cPersCod
        fg.TextMatrix(nItem, 9) = rs!cMovNro
        fg.TextMatrix(nItem, 10) = rs!nMovNro
        fg.TextMatrix(nItem, 11) = rs!nDocTpo
        fg.TextMatrix(nItem, 12) = rs!cDocNro
        fg.TextMatrix(nItem, 13) = rs!cCtaContCod
        fg.TextMatrix(nItem, 14) = GetFechaMov(rs!cMovNro, True)
        fg.TextMatrix(nItem, 23) = Format(rs!nPenalidad, gsFormatoNumeroView)
               
        If rs!nMovImporteSoles = rs!nmovimporte Then
            fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
        Else
            fg.TextMatrix(nItem, 15) = "0.00"
        End If
        fg.TextMatrix(nItem, 21) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
        fg.TextMatrix(nItem, 22) = Format(rs!MontoPagadoSUNATS, gsFormatoNumeroView)
        
        
        If (lsTipoInterfaz = "PAGO" Or lsTipoInterfaz = "lbReporte") Then
           fg.TextMatrix(nItem, 19) = IIf(rs!nMovImporteSoles = rs!nmovimporte, "SOLES", "DOLARES")
        End If
        fg.TextMatrix(nItem, 16) = IIf(IsNull(rs!cpersidnro), "RUC NO REGISTRADO", Trim(rs!cpersidnro))
     
    rs.MoveNext
Loop
RSClose rs
fg.Row = 1

oCon.CierraConexion

If Me.fg.TextMatrix(1, 3) = "" Then
    MsgBox " No existen datos para enviar a consulta SUNAT ", vbInformation, "Aviso"
    Me.cmdEnvio.Enabled = False
    Me.chkTodos.Enabled = False
Else
    Me.cmdEnvio.Enabled = True
    Me.chkTodos.Enabled = True
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fg_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim i As Integer
    
    Me.txtTot = 0
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            Me.txtTot = IIf(Me.txtTot = "0", 0, CCur(Me.txtTot)) + Me.fg.TextMatrix(i, 7)
        End If
    Next i
    Me.txtTot = Format(Me.txtTot, gsFormatoNumeroView)
End Sub

Private Sub Form_Activate()
    Me.txtFechaAl.SetFocus
End Sub

Private Sub Form_Load()
    Me.txtFechaAl = gdFecSis
    Me.txtFechaDel = DateAdd("m", -1, gdFecSis)
    Me.cmdEnvio.Enabled = False
    Me.chkTodos.Enabled = False
    
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFechaAl.SetFocus
    End If
End Sub
