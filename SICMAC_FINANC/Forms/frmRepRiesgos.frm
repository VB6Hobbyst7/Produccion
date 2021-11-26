VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepRiesgos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Movimientos por Lavado de Dinero"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmRepRiesgos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExceptuados 
      Cancel          =   -1  'True
      Caption         =   "&Exceptuados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5295
      TabIndex        =   12
      Top             =   2175
      Width           =   1140
   End
   Begin VB.CommandButton cmdReporteMensual 
      Caption         =   "&Mensual"
      Height          =   360
      Left            =   5295
      TabIndex        =   11
      Top             =   855
      Width           =   1140
   End
   Begin VB.CommandButton cmdResumen 
      Caption         =   "&Resumen"
      Height          =   360
      Left            =   5295
      TabIndex        =   10
      Top             =   1275
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   135
      Left            =   1425
      TabIndex        =   9
      Top             =   5025
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmRepRiesgos.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkCondenzado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Condensado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5070
      TabIndex        =   4
      Top             =   90
      Width           =   1410
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Todos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   525
      Width           =   1410
   End
   Begin MSMask.MaskEdBox mskIni 
      Height          =   285
      Left            =   615
      TabIndex        =   0
      Top             =   75
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5295
      TabIndex        =   6
      Top             =   4335
      Width           =   1140
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   5295
      TabIndex        =   5
      Top             =   1710
      Width           =   1140
   End
   Begin VB.ListBox lstRiesgos 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   15
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   840
      Width           =   5145
   End
   Begin MSMask.MaskEdBox mskFin 
      Height          =   285
      Left            =   2460
      TabIndex        =   1
      Top             =   90
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFin 
      Caption         =   "Fin :"
      Height          =   240
      Left            =   2055
      TabIndex        =   8
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lblInicio 
      Caption         =   "Inicio :"
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   105
      Width           =   660
   End
End
Attribute VB_Name = "frmRepRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gcCentralCom As String
Dim gcCentralPers As String

Private Sub chktodos_Click()
    Dim I As Integer
    
    For I = 0 To Me.lstRiesgos.ListCount - 1
        lstRiesgos.Selected(I) = IIf(chkTodos.value = 1, True, False)
    Next I
End Sub

Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lstRiesgos.SetFocus
    End If
End Sub

Private Sub cmdExceptuados_Click()
    Dim I As Integer
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim oPrevio As PrevioFinan.clsPrevioFinan
    Set oPrevio = New PrevioFinan.clsPrevioFinan
    Dim lsCodPers As String * 10
    Dim lsFecha As String * 8
    Dim lsNOmbre As String * 40
    Dim lsComenta As String * 40
    Dim lnPagina As Long
    Dim lnItem As Long
    
    If gbBitCentral Then
        sql = " Select PL.cPersCod cCodPers, Convert(DateTime, Left(PL.cMovNro,8)) dFecha, PE.cPersNombre cNomPers, PL.cComentario " _
            & " From PersExoLavDinero PL" _
            & " Inner Join Persona PE On PE.cPersCod = PL.cPersCod" _
            & " Where PL.nEstado  = 2 And PL.cMovNro = (Select Max(cMovNro) From PersExoLavDinero PA Where PL.cPersCod = PA.cPersCod)"
        oCon.AbreConexion
    Else
        sql = " Select PL.cCodPers, PL.dFecha, PE.cNomPers, PL.cComentario " _
            & " From DBPersona..PersLavado PL" _
            & " Inner Join DBPersona..Persona PE On PE.cCodPers = PL.cCodPers" _
            & " Where PL.cEstado  = '2' And PL.dFecha = (Select Max(dFecha) From PersLavado PA Where PL.cCodPers = PA.cCodPers)"
        oCon.AbreConexion 'Remota (Right(gsCodAge, 2))
    End If
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then

        lsCadena = lsCadena & CabeceraPagina("PERSONAS ECEPTUADAS DEL CONTRO DE LAV. DE DINERO", lnPagina, lnItem, "")
        lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Codigo;8; ;4;Fecha;6; ;4;Nombre;30; ;10;Comentario;20; ;20;", lnItem) & oImpresora.gPrnBoldOFF

        While Not rs.EOF
            lsCodPers = rs!cCodPers
            lsFecha = Format(rs!dFecha, gsFormatoFechaView)
            lsNOmbre = rs!cNomPers
            lsComenta = rs!cComentario
        
            lsCadena = lsCadena & lsCodPers & "  " & lsFecha & "  " & lsNOmbre & "  " & lsComenta & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1

            If lnItem > 54 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("PERSONAS ECEPTUADAS DEL CONTRO DE LAV. DE DINERO", lnPagina, lnItem, "")
                lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Codigo;8; ;4;Fecha;6; ;4;Nombre;30; ;10;Comentario;20; ;20;", lnItem) & oImpresora.gPrnBoldOFF
            End If
            rs.MoveNext
        Wend
    End If
    
    oCon.CierraConexion
    
    
    oPrevio.Show lsCadena, Caption, IIf(Me.chkCondenzado.value = 1, True, False), 64, gImpresora
    
End Sub

Private Sub cmdProcesar_Click()
    Dim I As Integer
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oPrevio As PrevioFinan.clsPrevioFinan
    Set oPrevio = New PrevioFinan.clsPrevioFinan
    
    If Not IsDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    ElseIf CDate(Me.mskFin.Text) < CDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar mayor que la fecha inicial.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    End If
    
    GetTipCambio CDate(Me.mskFin.Text)
    If Not gbBitCentral Then
        For I = 0 To Me.lstRiesgos.ListCount - 1
            If Me.lstRiesgos.Selected(I) Then
                If oCon.AbreConexion Then 'Remota(Right(Me.lstRiesgos.List(i), 2))
                    lsCadena = lsCadena & GetReporte(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), "1")
                    lsCadena = lsCadena & GetReporte(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), "2")
                    oCon.CierraConexion
                End If
            End If
        Next I
    Else
        oCon.AbreConexion
        For I = 0 To Me.lstRiesgos.ListCount - 1
            If Me.lstRiesgos.Selected(I) Then
                lsCadena = lsCadena & GetReporteCentral(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), "1", Right(Me.lstRiesgos.List(I), 2))
                lsCadena = lsCadena & GetReporteCentral(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), "2", Right(Me.lstRiesgos.List(I), 2))
            End If
        Next I
    End If
    
    GetTipCambio gdFecSis
    
    oPrevio.Show lsCadena, Caption, IIf(Me.chkCondenzado.value = 1, True, False), 64, gImpresora
End Sub

Private Sub cmdReporteMensual_Click()
    Dim I As Integer
    Dim lsCadena As String
    Dim oPrevio As PrevioFinan.clsPrevioFinan
    Set oPrevio = New PrevioFinan.clsPrevioFinan
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    If Not IsDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    ElseIf CDate(Me.mskFin.Text) < CDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar mayor que la fecha inicial.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    End If
    
    GetTipCambio CDate(Me.mskFin.Text)
    
    lsCadena = lsCadena & GetReporteMes(CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), "REPORTES DE PERSONAS MAYOR A US/. 50,000.00 o " & Format(oGen.GetParametro(2000, CaptacParametro.gMonMensLavDineroME) * gnTipCambio, "#,##0.00") & Me.mskIni.Text & " - " & Me.mskFin.Text)
    
    GetTipCambio gdFecSis
    
    oPrevio.Show lsCadena, Caption, IIf(Me.chkCondenzado.value = 1, True, False), 64, gImpresora
End Sub

Private Sub cmdResumen_Click()
    Dim I As Integer
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim oPrevio As PrevioFinan.clsPrevioFinan
    Set oPrevio = New PrevioFinan.clsPrevioFinan
    
    Dim lnTotSN As Currency
    Dim lnTotSM As Currency
    Dim lnTotDN As Currency
    Dim lnTotDM As Currency
    
    Dim lsAgencia As String * 45
    Dim lsMontoS As String * 15
    Dim lsMumeroS As String * 15
    Dim lsMontoD As String * 15
    Dim lsMumeroD As String * 15
    On Error GoTo ErrRepResumen
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If Not IsDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    ElseIf CDate(Me.mskFin.Text) < CDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar mayor que la fecha inicial.", vbInformation, "Aviso"
        mskFin.SetFocus
        Exit Sub
    End If
    
    GetTipCambio CDate(Me.mskFin.Text)
            
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("RESUMEN MOV LAVADO DE DINERO : " & Trim(gsNomCmac), lnPagina, lnItem, "NNN")
    lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Agencia;15; ;35;Num S/.;12;Mon S/.;10; ;13;Num $/.;8; ;3;Mon $/.;8; ;15;", lnItem) & oImpresora.gPrnBoldOFF
    
    If Not gbBitCentral Then
        For I = 0 To Me.lstRiesgos.ListCount - 1
            If Me.lstRiesgos.Selected(I) Then
                If oCon.AbreConexion Then 'Remota(Right(Me.lstRiesgos.List(i), 2))
                    lsCadena = lsCadena & GetReporteResumen(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), lnTotSN, lnTotSM, lnTotDN, lnTotDM)
                    oCon.CierraConexion
                End If
            End If
        Next I
    Else
        oCon.AbreConexion
        For I = 0 To Me.lstRiesgos.ListCount - 1
            If Me.lstRiesgos.Selected(I) Then
                lsCadena = lsCadena & GetReporteResumenCentral(oCon, CDate(Me.mskIni.Text), CDate(Me.mskFin.Text), Trim(Left(lstRiesgos.List(I), 50)), lnTotSN, lnTotSM, lnTotDN, lnTotDM, Right(Me.lstRiesgos.List(I), 2))
            End If
        Next I
        oCon.CierraConexion
    End If
    
    RSet lsMontoS = Format(lnTotSM, "#,##0.00")
    RSet lsMumeroS = Format(lnTotSN, "#,##0")
    RSet lsMontoD = Format(lnTotDM, "#,##0.00")
    RSet lsMumeroD = Format(lnTotDN, "#,##0")
    lsAgencia = ""
    lsCadena = lsCadena & String(120, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsAgencia & lsMumeroS & lsMontoS & lsMumeroD & lsMontoD & oImpresora.gPrnSaltoLinea
    
    GetTipCambio gdFecSis
    
    oPrevio.Show lsCadena, Caption, IIf(Me.chkCondenzado.value = 1, True, False), 64, gImpresora
Exit Sub
ErrRepResumen:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    gcCentralCom = "DBComunes.."
    gcCentralPers = "DBPersona.."
    
    sql = "select cAgeDescripcion + space(150) + cAgeCod Age from agencias "
    Set rs = oCon.CargaRecordSet(sql)
    
    Me.lstRiesgos.Clear

    While Not rs.EOF
        lstRiesgos.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
    CentraForm Me

    Set oCon = Nothing
    Set rs = Nothing

End Sub

Private Function GetReporte(pConec As DConecta, pdFecIni As Date, pdFecFin As Date, psTitulo As String, psMoneda As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsFecha As String * 21
    Dim lsCuenta As String * 12
    Dim lsMonto As String * 13
    Dim lsTrami As String * 35
    Dim lsOperacion As String * 29
    Dim lsCodCtaAnt As String
        
    sql = " Select LD.nNUmtran , LD.dFecTran,LD.cCodCta,LD.nMonto,LD.cProcedencia,LD.cCodPersTrami,PE.cNomPers TRAMI ,PET.cCodPers,PET.cNomPers, OP.cNomOpe From LavDinero LD" _
        & " Inner Join " & gcCentralPers & "Persona PE On LD.cCodPersTrami = PE.cCodPers" _
        & " Inner Join PersCuenta PC On LD.cCodCta = PC.cCodCta" _
        & " Inner Join " & gcCentralPers & "Persona PET On PC.cCodPers = PET.cCodPers" _
        & " Inner Join " & gcCentralCom & "Operacion OP On LD.cOpeCod = OP.cCodOpe" _
        & " Inner Join TransAho TA On TA.nNumTran = LD.nNumTran And TA.dFecTran = LD.dFecTran" _
        & " Where (TA.cFlag Is Null Or TA.cFlag In ('1','2'))" _
        & " And LD.dFecTran Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoFecha) & "' ANd SubString(LD.cCodCta,6,1) = '" & psMoneda & "' " _
        & " Order by LD.dFecTran"
    Set rs = pConec.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then

        lsCadena = lsCadena & CabeceraPagina("MOV LAVADO DE DINERO : " & psTitulo, lnPagina, lnItem, psMoneda)
        lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Fecha/Hora;15; ;10;Cuenta;6; ;10;Monto;6; ;9;Tramitador;15; ;10;Operación;18; ;20;Origen;6;", lnItem) & oImpresora.gPrnBoldOFF

        While Not rs.EOF
            If lsCodCtaAnt <> rs!nnumTran Then
                lsFecha = Format(rs!DFECTRAN, gsFormatoFechaHoraView)
                lsCuenta = rs!cCodCta
                RSet lsMonto = Format(rs!nMonto, "#,##0.00")
                lsTrami = rs!TRAMI
                lsOperacion = rs!cNomOpe

                lsCadena = lsCadena & lsFecha & " " & lsCuenta & " " & lsMonto & "  " & lsTrami & " " & lsOperacion & " " & rs!cProcedencia & oImpresora.gPrnSaltoLinea
                lnItem = lnItem + 1

                If lnItem > 54 Then
                    lnItem = 0
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    lsCadena = lsCadena & CabeceraPagina("MOV LAVADO DE DINERO : " & psTitulo, lnPagina, lnItem, psMoneda)
                    lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Fecha/Hora;15; ;10;Cuenta;6; ;10;Monto;6; ;9;Tramitador;15; ;10;Operación;18; ;20;Origen;6;", lnItem) & oImpresora.gPrnBoldOFF
                End If
            End If
            lsCodCtaAnt = rs!nnumTran
            rs.MoveNext
        Wend
    End If


    If lsCadena <> "" Then
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If

    GetReporte = lsCadena
End Function

Private Function GetReporteMes(pdFecIni As Date, pdFecFin As Date, psTitulo As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsFecha As String * 50
    Dim lsCuenta As String * 12
    Dim lsMonto As String * 20
    Dim lsTrami As String * 35
    Dim lsOperacion As String * 35
    Dim lsCodCtaAnt As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oConLocal As DConecta
    Set oConLocal = New DConecta
    Dim lsCadenaConLocal As String
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    Dim sqlAdd As String
    
    oCon.AbreConexion 'Remota "07", , , "03"
    oConLocal.AbreConexion
    'lsCadenaConLocal = oConLocal.StringServidorRemoto("07", "04")
    
    If Not gbBitCentral Then
        sql = " Select Sum(TRANS.nMonTran) SUMA, PERS.cCodPers, PERS.cNomPers, PERS.cDirPers From " _
            & "     (Select PC.cCodCta, PE.cCodPers, PE.cNomPers, PE.cDirPers from PersCuentaConsol PC" _
            & "              Inner Join " & gcCentralPers & "Persona PE On PC.cCodPers = PE.cCodPers" _
            & "              Where cRelaCta = 'TI'" _
            & "     ) As PERS" _
            & " Inner Join" _
            & "     ( Select distinct dFectran, TA.cCodCta, Case Substring(TA.cCodCta,6,1) WHen '1' Then TA.nMonTran When '2' Then TA.nMonTran * " & gnTipCambio & " End nMonTran  from TransAhoConsol TA" _
            & "         Inner Join" _
            & "             (Select cCodOpe From " & gcCentralCom & "Operacion OPE" _
            & "                 Inner Join (Select cGruProd from " & gcCentralCom & "OpeGru where Substring(cGruProd,3,3) Between '000' And '300' And cIngEgr = 'I') As OG On OPE.cGruProd = OG.cGruProd) As OPE On TA.cCodOpe = OPE.cCodOpe" _
            & "                    Where (cFlag is Null Or cFlag In ('1','2')) And (TA.dFecTran Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoFecha) & "')) As TRANS" _
            & "           On PERS.cCodCta = TRANS.cCodCta Where nMonTran < (" & oGen.GetParametro(2000, CaptacParametro.gMonOpeLavDineroME) & " * " & gnTipCambio & ")" _
            & "         Group by PERS.cCodPers, PERS.cNomPers, PERS.cDirPers" _
            & "     Having Sum(TRANS.nMonTran) >= (" & oGen.GetParametro(2000, CaptacParametro.gMonMensLavDineroME) & " * " & gnTipCambio & ")" _
            & "     And PERS.cCodPers Not In" _
            & "       ( Select cCodPers from [128.107.2.2].dbcmact07.dbo.perslavado A" _
            & "         Where A.dFecha = (Select Top 1 B.dFecha from [128.107.2.2].dbcmact07.dbo.perslavado B Where A.cCodPers = B.cCodPers Order by B.dFecha desc)" _
            & "         And cEstado = '2')" _
            & " Order by SUMA Desc"
        Set rs = oCon.CargaRecordSet(sql)
    
    
    Else
        sqlAdd = " Select distinct M.cMovNro dFectran, MC.cCtaCod , Case Substring(MC.cCtaCod,9,1) WHen '" & Moneda.gMonedaNacional & "' Then MC.nMonto When '" & Moneda.gMonedaExtranjera & "' Then MC.nMonto * " & gnTipCambio & " End nMonTran" _
               & "     From Mov M" _
               & "     Inner Join MovCap MC On M.nMovNro = MC.nMovNro" _
               & "     Inner Join" _
               & "         (Select cOpeCod From GruposOpe OPE" _
               & "                Inner Join" _
               & "             (Select cGrupoCod From GrupoOpe" _
               & "            Where Substring(cGrupoCod,3,3) Between '000' And '300' And cIngEgr In ('I','E'))" _
               & "            As OG On OPE.cGrupoCod = OG.cGrupoCod)" _
               & "     As OPE On MC.cOpeCod = OPE.cOpeCod" _
               & "     Where (M.nMovFlag = 0) And (Left(M.cMovNro,8) Between '" & Format(pdFecIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "') UNION ALL " _
               & " Select distinct M.cMovNro dFectran, MC.cCtaCod , Case Substring(MC.cCtaCod,9,1) WHen '" & Moneda.gMonedaNacional & "' Then MC.nMonto When '" & Moneda.gMonedaExtranjera & "' Then MC.nMonto * " & gnTipCambio & " End nMonTran" _
               & "     From Mov M" _
               & "     Inner Join MovCol MC On M.nMovNro = MC.nMovNro" _
               & "     Inner Join" _
               & "         (Select cOpeCod From GruposOpe OPE" _
               & "                Inner Join" _
               & "             (Select cGrupoCod From GrupoOpe" _
               & "            Where Substring(cGrupoCod,3,3) Between '301' And '999' And cIngEgr In ('I','E'))" _
               & "            As OG On OPE.cGrupoCod = OG.cGrupoCod)" _
               & "     As OPE On MC.cOpeCod = OPE.cOpeCod" _
               & "     Where (M.nMovFlag = 0) And (Left(M.cMovNro,8) Between '" & Format(pdFecIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "')"
            
        sql = " Select Sum(TRANS.nMonTran) SUMA, PERS.cPersCod, PERS.cPersNombre, PERS.cPersDireccDomicilio " _
            & " From" _
            & " (Select PP.cCtaCod, PE.cPersCod, PE.cPersNombre, PE.cPersDireccDomicilio from ProductoPersona PP" _
            & "     Inner Join Persona PE On PP.cPersCod = PE.cPersCod" _
            & "     Where PP.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & ") As PERS" _
            & "     Inner Join" _
            & "     (" & sqlAdd & ") As TRANS On PERS.cCtaCod = TRANS.cCtaCod Where nMonTran < (" & oGen.GetParametro(2000, CaptacParametro.gMonOpeLavDineroME) & " * " & gnTipCambio & ") Group by PERS.cPersCod, PERS.cPersNombre, PERS.cPersDireccDomicilio" _
            & "     Having Sum(TRANS.nMonTran) >= (" & oGen.GetParametro(2000, CaptacParametro.gMonMensLavDineroME) & " * " & gnTipCambio & ")     And PERS.cPersCod Not In" _
            & "     ( Select A.cPersCod From PersExoLavDinero A Where A.cMovNro = (Select Top 1 B.cMovNro From PersExoLavDinero B Where A.cPersCod = B.cPersCod Order by B.cMovNro Desc) And nEstado = '2')" _
            & "     Order by SUMA Desc"
        
            Set rs = oConLocal.CargaRecordSet(sql)
    End If
    

    If Not (rs.EOF And rs.BOF) Then

        lsCadena = lsCadena & CabeceraPagina("MOV X MES: " & psTitulo, lnPagina, lnItem, "NNN")
        lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Persona;15; ;40;Monto;15; ;10;Direccion;15; ;30;", lnItem) & oImpresora.gPrnBoldOFF

        While Not rs.EOF
            lsFecha = Format(rs!cPersNombre, gsFormatoFechaHoraView)
            RSet lsMonto = Format(rs!SUMA, "#,##0.00")
            lsOperacion = rs!cPersDireccDomicilio

            lsCadena = lsCadena & lsFecha & " " & lsMonto & "       " & lsOperacion & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1

            If lnItem > 54 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("MOV X MES: " & psTitulo, lnPagina, lnItem, "NNN")
                lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Persona;15; ;40;Monto;10; ;15;Direccion;15; ;30;", lnItem) & oImpresora.gPrnBoldOFF
            End If
            rs.MoveNext
        Wend
    End If


    If lsCadena <> "" Then
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If

    GetReporteMes = lsCadena
End Function

Private Function GetReporteResumen(pConec As DConecta, pdFecIni As Date, pdFecFin As Date, psAgencia As String, pnTotSN As Currency, pnTotSM As Currency, pnTotDN As Currency, pnTotDM As Currency) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsAgencia As String * 45
    Dim lsMontoS As String * 15
    Dim lsMumeroS As String * 15
    Dim lsMontoD As String * 15
    Dim lsMumeroD As String * 15

    sql = " Select Substring(TA.cCodCta,6,1) Moneda, Count(*) Num, sum (LD.nMonto) Monto From LavDinero LD" _
        & " Inner Join " & gcCentralCom & "Operacion OP On LD.cOpeCod = OP.cCodOpe" _
        & " Inner Join TransAho TA On TA.nNumTran = LD.nNumTran And TA.dFecTran = LD.dFecTran" _
        & " Where (TA.cFlag Is Null Or TA.cFlag In ('1','2'))" _
        & " And LD.dFecTran Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoFecha) & "'" _
        & " Group by Substring(TA.cCodCta,6,1)"
    Set rs = pConec.CargaRecordSet(sql)

    RSet lsMontoS = Format(0, "#,##0.00")
    RSet lsMumeroS = Format(0, "#,##0")
    RSet lsMontoD = Format(0, "#,##0.00")
    RSet lsMumeroD = Format(0, "#,##0")

    If Not (rs.EOF And rs.BOF) Then
        lsAgencia = psAgencia
        While Not rs.EOF
            If rs!Moneda = "1" Then
                RSet lsMontoS = Format(rs!Monto, "#,##0.00")
                RSet lsMumeroS = rs!Num
                pnTotSN = pnTotSN + rs!Num
                pnTotSM = pnTotSM + rs!Monto
            Else
                RSet lsMontoD = Format(rs!Monto, "#,##0.00")
                RSet lsMumeroD = rs!Num
                pnTotDN = pnTotDN + rs!Num
                pnTotDM = pnTotDM + rs!Monto
            End If
            rs.MoveNext
        Wend

        lsCadena = lsCadena & lsAgencia & lsMumeroS & lsMontoS & lsMumeroD & lsMontoD & oImpresora.gPrnSaltoLinea

    End If

    GetReporteResumen = lsCadena
End Function

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub


Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkTodos.SetFocus
    End If
End Sub

Private Function CabeceraPagina(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, Optional psMoneda As String = "1") As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lsCadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lsCadena = ""

    lsC1 = oImpresora.gPrnNegritaON & Format(gdFecSis, "dd/mm/yyyy") & oImpresora.gPrnNegritaOFF
    lsC2 = oImpresora.gPrnNegritaON & Format(Time, "hh:mm:ss AMPM") & oImpresora.gPrnNegritaOFF
    lsC3 = oImpresora.gPrnNegritaON & "PAGINA Nro. " & Str(pnPagina) & oImpresora.gPrnNegritaOFF
    lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    lsCadena = oImpresora.gPrnNegritaON & lsCadena & "CMACT" & Space(39 - Len(lsC3) + 10 - Len("CMACT")) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
  
    If psMoneda = "1" Then
        lsCadena = lsCadena & oImpresora.gPrnNegritaON & gsNomAge & "-Soles" + Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(gsNomAge)) & lsC2 & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    ElseIf psMoneda = "NNN" Then
        lsCadena = lsCadena & oImpresora.gPrnNegritaON & gsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(gsNomAge)) & lsC2 & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & oImpresora.gPrnNegritaON & gsNomAge & "-Dolares" + Space(112 - Len(lsC2) - Len(lsC2) + 10 - Len(gsNomAge)) & lsC2 & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    End If
    lsCadena = lsCadena & oImpresora.gPrnNegritaON & CentrarCadena(psTitulo, 104) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    
    CabeceraPagina = lsCadena
End Function

Private Function Encabezado(psCadena As String, pnItem As Long, Optional pnMargen As Integer = 0) As String
    Dim lsCadena As String
    Dim lsCampo As String
    Dim lnLonCampo As Long
    Dim lnTotalLinea As Long
    Dim lnPos As Long
    Dim lsResultado As String
    Dim I As Long
    Dim lsLineas As String
    
    lsResultado = ""
    lnTotalLinea = 0
        
    lsCadena = psCadena
    pnItem = pnItem + 3
    
    While lsCadena <> ""
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        lsCampo = Left(lsCadena, lnPos - 1)
        lsCadena = Mid(lsCadena, lnPos + 1)
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        
        lnLonCampo = CCur(Left(lsCadena, lnPos - 1))
        
        lsCadena = Mid(lsCadena, lnPos + 1)
        
        lnTotalLinea = lnTotalLinea + lnLonCampo
        
        lsResultado = lsResultado + Space(lnLonCampo - Len(lsCampo)) & lsCampo
    Wend
        
    lsResultado = Space(pnMargen) & lsResultado & oImpresora.gPrnSaltoLinea
    
    lsLineas = Space(pnMargen) & String(lnTotalLinea + 1, "-") & oImpresora.gPrnSaltoLinea
    
    lsResultado = lsLineas & lsResultado & lsLineas
    
    Encabezado = lsResultado
End Function

Private Function GetReporteCentral(pConec As DConecta, pdFecIni As Date, pdFecFin As Date, psTitulo As String, psMoneda As String, psAgencia As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsFecha As String * 21
    Dim lsCuenta As String * 12
    Dim lsMonto As String * 13
    Dim lsTrami As String * 35
    Dim lsOperacion As String * 29
    Dim lsCodCtaAnt As String
    Dim lnCorr As Long
    
    sql = " Select LD.nMovNro nNumTran , dbo.FechaHoraMov(M.cMovNro) dFecTran, MCAP.cCtaCod cCodCta, Sum(MCAP.nMonto) nMonto, Substring(M.cMovNro,18,2) cProcedencia, LD.cPersCod cCodPersTrami, PE.cPersNombre TRAMI ,PET.cPersCod cCodPers, PET.cPersNombre cNomPers, OP.cOpeDesc cNomOpe " _
        & " From MovLavDinero LD" _
        & " Inner Join Mov  M On M.nMovNro = LD.nMovNro" _
        & " Inner Join MovCap MCAP On M.nMovNro = MCAP.nMovNro" _
        & " Inner Join Persona PE On LD.cPersCod = PE.cPersCod" _
        & " Inner Join ProductoPersona PC On MCAP.cCtaCod = PC.cCtaCod" _
        & " Inner Join Persona PET On PC.cPersCod = PET.cPersCod" _
        & " Inner Join OpeTpo OP On M.cOpeCod = OP.cOpeCod" _
        & " Where (M.nMovFlag = " & MovFlag.gMovFlagVigente & ") And M.cMovNro Between '" & Format(pdFecIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "'" _
        & " And SubString(MCAP.cCtaCod,9,1) = '" & psMoneda & "' And Substring(M.cMovNro,18,2) = '" & psAgencia & "'" _
        & " Group By LD.nMovNro, dbo.FechaHoraMov(M.cMovNro), MCAP.cCtaCod , Substring(M.cMovNro,18,2) , LD.cPersCod , PE.cPersNombre  ,PET.cPersCod , PET.cPersNombre , OP.cOpeDesc" _
        & " Order by dFecTran"
    Set rs = pConec.CargaRecordSet(sql)
     
    If Not (rs.EOF And rs.BOF) Then

        lsCadena = lsCadena & CabeceraPagina("MOV LAVADO DE DINERO : " & psTitulo, lnPagina, lnItem, psMoneda)
        lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Item;4; ;2;Fecha/Hora;15; ;10;Cuenta;6; ;10;Monto;6; ;9;Tramitador;15; ;10;Operación;18; ;20;Origen;6;", lnItem) & oImpresora.gPrnBoldOFF
        lnCorr = 0
        While Not rs.EOF
            If lsCodCtaAnt <> rs!nnumTran Then
                lsFecha = Format(rs!DFECTRAN, gsFormatoFechaHoraView)
                lsCuenta = rs!cCodCta
                RSet lsMonto = Format(rs!nMonto, "#,##0.00")
                lsTrami = rs!TRAMI
                lsOperacion = rs!cNomOpe
                lnCorr = lnCorr + 1
                lsCadena = lsCadena & Format(lnCorr, "0000      ") & lsFecha & " " & lsCuenta & " " & lsMonto & "  " & lsTrami & " " & lsOperacion & " " & rs!cProcedencia & oImpresora.gPrnSaltoLinea
                lnItem = lnItem + 1

                If lnItem > 54 Then
                    lnItem = 0
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    lsCadena = lsCadena & CabeceraPagina("MOV LAVADO DE DINERO : " & psTitulo, lnPagina, lnItem, psMoneda)
                    lsCadena = lsCadena & oImpresora.gPrnBoldON & Encabezado("Fecha/Hora;15; ;10;Cuenta;6; ;10;Monto;6; ;9;Tramitador;15; ;10;Operación;18; ;20;Origen;6;", lnItem) & oImpresora.gPrnBoldOFF
                End If
            End If
            lsCodCtaAnt = rs!nnumTran
            rs.MoveNext
        Wend
    End If


    If lsCadena <> "" Then
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
    End If

    GetReporteCentral = lsCadena
End Function

Private Function GetReporteResumenCentral(pConec As DConecta, pdFecIni As Date, pdFecFin As Date, psAgencia As String, pnTotSN As Currency, pnTotSM As Currency, pnTotDN As Currency, pnTotDM As Currency, psAgenciaCod As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsAgencia As String * 45
    Dim lsMontoS As String * 15
    Dim lsMumeroS As String * 15
    Dim lsMontoD As String * 15
    Dim lsMumeroD As String * 15

    sql = " Select SubString(MCAP.cCtaCod,9,1) Moneda, Count(Distinct M.nMovNro) Num, Sum(MCAP.nMonto) Monto " _
        & " From MovLavDinero LD" _
        & " Inner Join Mov  M On M.nMovNro = LD.nMovNro" _
        & " Inner Join MovCap MCAP On M.nMovNro = MCAP.nMovNro" _
        & " Inner Join Persona PE On LD.cPersCod = PE.cPersCod" _
        & " Inner Join ProductoPersona PC On MCAP.cCtaCod = PC.cCtaCod" _
        & " Inner Join Persona PET On PC.cPersCod = PET.cPersCod" _
        & " Inner Join OpeTpo OP On M.cOpeCod = OP.cOpeCod" _
        & " Where (M.nMovFlag = " & MovFlag.gMovFlagVigente & ") And M.cMovNro Between '" & Format(pdFecIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFecFin), gsFormatoMovFecha) & "'" _
        & " And Substring(M.cMovNro,18,2) = '" & psAgenciaCod & "'" _
        & " Group By SubString(MCAP.cCtaCod,9,1)" _
        & " "
    Set rs = pConec.CargaRecordSet(sql)

    RSet lsMontoS = Format(0, "#,##0.00")
    RSet lsMumeroS = Format(0, "#,##0")
    RSet lsMontoD = Format(0, "#,##0.00")
    RSet lsMumeroD = Format(0, "#,##0")

    If Not (rs.EOF And rs.BOF) Then
        lsAgencia = psAgencia
        While Not rs.EOF
            If rs!Moneda = "1" Then
                RSet lsMontoS = Format(rs!Monto, "#,##0.00")
                RSet lsMumeroS = rs!Num
                pnTotSN = pnTotSN + rs!Num
                pnTotSM = pnTotSM + rs!Monto
            Else
                RSet lsMontoD = Format(rs!Monto, "#,##0.00")
                RSet lsMumeroD = rs!Num
                pnTotDN = pnTotDN + rs!Num
                pnTotDM = pnTotDM + rs!Monto
            End If
            rs.MoveNext
        Wend

        lsCadena = lsCadena & lsAgencia & lsMumeroS & lsMontoS & lsMumeroD & lsMontoD & oImpresora.gPrnSaltoLinea

    End If

    GetReporteResumenCentral = lsCadena
End Function


