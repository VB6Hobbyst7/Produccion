VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmArchRecibosHonorarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Recibos por Honorarios"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   ForeColor       =   &H8000000F&
   Icon            =   "frmArchRecibosHonorarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   3495
      Begin VB.OptionButton opPrestadores 
         Caption         =   "Prestadores de Servicios 4ta Cat (*.ps4)"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   3135
      End
      Begin VB.OptionButton optDatos 
         Caption         =   "Datos Persosnales 4ta Categoria"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Detalle del comprobante (*.4ta)"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   210
      Left            =   795
      TabIndex        =   6
      Top             =   2700
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir Observaciones"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   2340
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   2025
      TabIndex        =   3
      Top             =   1875
      Width           =   1140
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   1875
      Width           =   1080
   End
   Begin MSMask.MaskEdBox MskAl 
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   165
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   -240
      TabIndex        =   5
      Top             =   -225
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmArchRecibosHonorarios.frx":08CA
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
   Begin VB.Label Label2 
      Caption         =   "Al :"
      Height          =   300
      Left            =   705
      TabIndex        =   2
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "frmArchRecibosHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTmp As New ADODB.Recordset
Dim sTDOC As String * 2 '13
'Dim sNDOC As String * 8 '14
Dim sNDOC As String * 15 'EJVG20121109
Dim sP As String * 1
Dim sql As String
Dim j As Integer
Dim lsArc As String
Dim sLinea As String
Dim NumeroArchivo As Integer
'**** Creado ALPA
'**** 18/02/2008
Private Sub darReporteCuartaDatos(Optional Opcion As Integer)
Dim rs As New ADODB.Recordset
Dim sRUC  As String * 11 '1
Dim sApPaterno  As String * 20 '2
Dim sApMaterno As String * 20 '3
Dim sNombres  As String * 20 '4
Dim sFechaNac  As String * 10 '7
Dim sSexo  As String * 1 '7
Dim sNacionalidad  As String * 4 '7
Dim sTelefono  As String * 15 '7
Dim sRUCTemp As String
Dim sTipoCom As String * 1 '8
Dim sNroSerie As String * 4 '5
Dim sNroRecibo As String * 8 '6
Dim sFecEmision As String * 10 '7
Dim sFecPago As String * 10 '7
Dim sMonto As String * 15 '8
Dim sRetencion As String * 1 '9
Dim sNombreTemp As String
Dim oConConsol As DConecta
Set oConConsol = New DConecta
Dim i As Integer
Dim lsNroSerie As String
Dim iNroSerie As Integer

On Error GoTo ArchivoErr

j = 0
sP = "|"
PB.Visible = True
PB.value = 0
PB.Min = 0

If rsTmp.State = adStateOpen Then
    rsTmp.Close
End If
rsTmp.Fields.Append "CodPers", adChar, 15
rsTmp.Fields.Append "NomPers", adChar, 50
rsTmp.Open
oConConsol.AbreConexion
NumeroArchivo = FreeFile


If Opcion = 3 Then
lsArc = App.path & "\Spooler\0601" & Format(Me.MskAl, "YYYYMM") & gsRUC & ".4ta"
sql = "stp_sel_FinRentaCuarta_4ta '" & CStr(Year(MskAl)) & "', '" & Format(CStr(Month(MskAl)), "00") & "'"
End If

Set rs = oConConsol.CargaRecordSet(sql)
Set rs.ActiveConnection = Nothing
If rs.RecordCount > 0 Then
    PB.Max = rs.RecordCount
    PB.Visible = False
End If

If Not (rs.EOF And rs.BOF) Then
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    While Not rs.EOF
        j = rs.RecordCount

        'If IsNull(rs!RUC) And rs!cPersNombre <> sNombres Then
        If rs!NroDocId = "" And rs!cPersNombre <> sNombres Then 'EJVG20121109
            rsTmp.AddNew
            rsTmp!CodPers = rs!cPersCod
            'rsTmp!NomPers = rs!cApePat & Space(1) & rs!cApeMat & Space(1) & rs!cPrimerNombre
            rsTmp!NomPers = rs!cPersNombre 'EJVG20121109
        End If
        
        sApPaterno = PstaNombre(rs!cPersNombre, True)


        i = InStr(1, rs!cPersNombre, "/", vbTextCompare)
        If i = 0 Then
            MsgBox "Corregir Tipo de Documento de: " & sApPaterno & space(2) & sApMaterno, vbInformation, "Aviso"
            PB.value = 0
            PB.Visible = False
            cmdGenerar.Enabled = False
            Close #NumeroArchivo   ' Cierra el archivo.
            oConConsol.CierraConexion
            Set rsTmp = Nothing
            Exit Sub
        Else
            sApPaterno = Trim(Mid(rs!cPersNombre, 1, InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1))
        End If


        If InStr(1, rs!cPersNombre, "\", vbTextCompare) <> 0 Then
            sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "\", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "\", vbTextCompare) - 1)
        
        Else
            sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "/", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1)
        End If
        sNombres = LTrim(Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, ",", vbTextCompare) + 1, 100))
            sTDOC = rs!TDOC
            sNDOC = IIf(IsNull(rs!NroDocId), "", rs!NroDocId)
        
        If Opcion = 3 Then
        sTipoCom = rs!TCOMPRO
        sNroSerie = Format(rs!Serie, "0000")
        'EJVG20121207 ***
        lsNroSerie = ""
        
        'TORE - RFC
        If Not IsNumeric(sNroSerie) Then
            lsNroSerie = sNroSerie
        Else
            For iNroSerie = 1 To Len(sNroSerie)
                If Not IsNumeric(Mid(sNroSerie, iNroSerie, 1)) Then
                    lsNroSerie = Val(lsNroSerie & "0")
                Else
                    lsNroSerie = Val(lsNroSerie & Mid(sNroSerie, iNroSerie, 1))
                End If
            Next
        End If
        'END TORE
        
        'TORE - Comentado
'         For iNroSerie = 1 To Len(sNroSerie)
'            If Not IsNumeric(Mid(sNroSerie, iNroSerie, 1)) Then
'                lsNroSerie = lsNroSerie & "0"
'            Else
'                lsNroSerie = lsNroSerie & Mid(sNroSerie, iNroSerie, 1)
'            End If
'        Next

        'END EJVG *******
        sNroRecibo = Right(rs!nro, 8)
        sMonto = Format(rs!Bruto, "####0.00")
        sFecEmision = Format(rs!dDocFecha, "dd/mm/yyyy")
        sFecPago = Format(rs!FechaPago, "dd/mm/yyyy")
        sRetencion = IIf(rs!IMP = 0, "0", "1")
        'Inicio - Modificado por ORCR 03/02/2014
        'sLinea = Trim(sTDOC) & sP & Trim(sNDOC) & sP & Trim(sTipoCom) & sP & Val(lsNroSerie) & sP & Trim(sNroRecibo) & sP & Trim(sMonto) & sP & Trim(sFecEmision) & sP & Trim(sFecPago) & sP & Trim(sRetencion) & sP
        'sLinea = Trim(sTDOC) & sP & Trim(sNDOC) & sP & Trim(sTipoCom) & sP & Val(lsNroSerie) & sP & Trim(sNroRecibo) & sP & Trim(sMonto) & sP & Trim(sFecEmision) & sP & Trim(sFecPago) & sP & Trim(sRetencion) & sP & "3" & sP + sP
        sLinea = Trim(sTDOC) & sP & Trim(sNDOC) & sP & Trim(sTipoCom) & sP & lsNroSerie & sP & Trim(sNroRecibo) & sP & Trim(sMonto) & sP & Trim(sFecEmision) & sP & Trim(sFecPago) & sP & Trim(sRetencion) & sP & "" & sP + sP 'TORE - RFC
        'Fin - Modificado por ORCR 03/02/2014
        End If
        
        
        
        '/*************/
        Print #NumeroArchivo, sLinea

        PB.value = PB.value + 1
        rs.MoveNext
    Wend
    PB.value = 0
    PB.Visible = False
    cmdGenerar.Enabled = False
    Close #NumeroArchivo   ' Cierra el archivo.
    oConConsol.CierraConexion
    MsgBox "Archivo Generado en su Spooler", vbInformation, "Aviso"
    If RSVacio(rsTmp) Then Exit Sub
'    If rsTmp.RecordCount > 0 Then
'        MsgBox "El archivo presenta observaciones", vbInformation, "Aviso"
'        cmdImprimir.Enabled = True
'        cmdImprimir.SetFocus
'    End If
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
    PB.Visible = False
    If rsTmp.State = adStateOpen Then
        rsTmp.Close
    End If
    Set rsTmp = Nothing

End If

Exit Sub
ArchivoErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    PB.Visible = False
    cmdGenerar.Enabled = False
    Close #NumeroArchivo   ' Cierra el archivo.
    oConConsol.CierraConexion
End Sub
'******
Private Sub cmdGenerar_Click()
'*** Modificado ALPA
'*** Fecha 18/02/2008

If optDetalle.value Then
    Call darReporteCuartaDatos(3)
End If
'EJVG20121109 ***
If opPrestadores.value Then
    Call darReporteCuarta_PrestadoresServicio
End If
'END EJVG *******
'******Comentado por ALPA
'******16/02/2008
'Dim J As Integer
'Dim lsArc As String
'Dim sP As String * 1
'Dim SQL As String
'Dim NumeroArchivo As Integer
'Dim sLinea As String
'Dim rs As New ADODB.Recordset
'Dim sRUC  As String * 11 '1
'Dim sApPaterno  As String * 20 '2
'Dim sApMaterno As String * 20 '3
'Dim sNombres  As String * 20 '4
'Dim sNroSerie As String * 3 '5
'Dim sNroRecibo As String * 8 '6
'Dim sFecEmision As String * 10 '7
'Dim sMonto As String * 15 '8
'Dim sRetencion As String * 1 '9
'Dim sPorcRet As String  '10 ---- 10 %
'Dim sRetIES As String * 1 '11
'Dim sNroOrden As String * 1 '12
'Dim sTDOC As String * 2 '13
'Dim sNDOC As String * 8 '14
'Dim sRUCTemp As String
'Dim sNombreTemp As String
'Dim oConConsol As DConecta
'Set oConConsol = New DConecta
'Dim I As Integer
'
'
'On Error GoTo ArchivoErr
'
'J = 0
'sP = "|"
'PB.Visible = True
'PB.value = 0
'PB.Min = 0
'
'If rsTmp.State = adStateOpen Then
'    rsTmp.Close
'End If
'rsTmp.Fields.Append "CodPers", adChar, 15
'rsTmp.Fields.Append "NomPers", adChar, 50
'rsTmp.Open
'oConConsol.AbreConexion
''oConConsol.AbreConexionRemota "07", , , "03"
'NumeroArchivo = FreeFile
''**Modificado por ALPA 20080214 ***********************
''lsArc = App.path & "\Spooler\0621" & gsRUC & Format(Me.MskAl, "YYYYMM") & ".txt"
'lsArc = App.path & "\Spooler\" & gsRUC & ".t03"
''******************************************************
'
''''Sql = ""
''''Sql = "Select d.cPersCod, a.nmovnro,cPersNombre, dDocFecha,  " & _
''''        "(Select cPersIDnro from dbCmactAux..persid pid where pid.cPersCod = d.cperscod And cPersIDTpo = 2) RUC,  " & _
''''        "left(cDocNro,PATINDEX('%-%',cDocNro) - 1) Serie, substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) Nro,  " & _
''''        "(Select sum(nMovImporte) from dbCmactAux..movcta mcta where mcta.nmovnro = a.nmovnro and nMovImporte > 0) Bruto,  " & _
''''        "(Select abs(nMovImporte) from dbCmactAux..movcta mcta where mcta.nmovnro = a.nmovnro and nMovImporte < 0 And cctacontcod = '2114020201' ) IMP  " & _
''''"from dbCmactAux..mov a  " & _
''''"inner join dbCmactAux..movdoc b on a.nmovnro = b.nmovnro " & _
''''"inner join dbCmactAux..movgasto c on a.nmovnro = c.nmovnro  " & _
''''"inner join dbCmactAux..persona d on c.cperscod = d.cperscod  " & _
''''"where datediff(m, dDocFecha, '" & Format(MskAl.Text, "yyyy/mm/dd") & "') = 0  and ndoctpo = 2 and nmovflag = 0 and nMovEstado = 10 order by cPersNombre"
'
''GCZA
''SQL = ""
''SQL = SQL & " select p.cPersCod,pei.cPersIDnro RUC,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) Serie, substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) Nro, ddocfecha,"
''SQL = SQL & " IsNull(sum(case when nmovimporte>0 then nmovimporte end),0) Bruto,"
''SQL = SQL & " IsNull(sum(case when nMovImporte < 0 And cctacontcod = '2114020201' then nmovimporte end),0) IMP "
''SQL = SQL & " from mov m"
''SQL = SQL & " Inner join MovCta mc on m.nmovnro=mc.nmovnro"
''SQL = SQL & " Inner Join OpeTpo o on m.copecod=o.copecod"
''SQL = SQL & " Inner join movdoc md on md.nmovnro=m.nmovnro and ndoctpo=2"
''SQL = SQL & " Inner Join MovGasto mg on mg.nmovnro=m.nMovNro"
''SQL = SQL & " Inner Join Persona p on p.cPersCod=mg.cPersCod"
''SQL = SQL & " Left Join PersID pei on pei.cPersCod=p.cPersCod and cPersIDTpo=2"
''SQL = SQL & " Where m.nmovflag = 0 And nmovestado = 10"
''SQL = SQL & " and cmovnro like '" & Format(MskAl.Text, "yyyymm") & "%'"
''SQL = SQL & " and m.nmovnro not in (select nmovnroref from MovRef mr Inner Join Mov m on mr.nMovNro=mr.nMovNro where nmovflag=0 and (cagecodref='' or cagecodref is null)  )"
''SQL = SQL & " and m.copecod not like '7%'"
''SQL = SQL & " group by p.cPersCod,m.nmovnro, pei.cPersIDnro,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) , substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) , ddocfecha"
''SQL = SQL & " Union All"
''SQL = SQL & " select p.cPersCod,pei.cPersIDnro RUC,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) Serie, substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) Nro, ddocfecha,"
''SQL = SQL & " IsNull(sum(case when nmovimporte>0 then nmovimporte end),0) Bruto,"
''SQL = SQL & " IsNull(sum(case when nMovImporte < 0 And cctacontcod = '2114020201' then isnull(nmovimporte,0) end),0) IMP"
''SQL = SQL & " from mov m"
''SQL = SQL & " Inner Join MovRef mr on mr.nMovNro=m.nMovNro and nmovflag=0 and isnull(cagecodref,'')=''"
''SQL = SQL & " Inner Join OpeTpo o on m.copecod=o.copecod"
''SQL = SQL & " Inner Join Mov M1 on M1.nMovNro=mr.nMovNroRef"
''SQL = SQL & " Inner join MovCta mc on m1.nmovnro=mc.nmovnro"
''SQL = SQL & " Inner join movdoc md on md.nmovnro=m1.nmovnro and ndoctpo=2"
''SQL = SQL & " Inner Join MovGasto mg on mg.nmovnro=m.nMovNro"
''SQL = SQL & " Inner Join Persona p on p.cPersCod=mg.cPersCod"
''SQL = SQL & " Left Join PersID pei on pei.cPersCod=p.cPersCod and cPersIDTpo=2"
''SQL = SQL & " Where m.nmovflag = 0 And m.nmovestado = 10 And m1.nmovflag = 0  and m.copecod <> '421116' "
''SQL = SQL & " and m.cmovnro like '" & Format(MskAl.Text, "yyyymm") & "%'"
''SQL = SQL & " and m.nmovnro not in (select nmovnroref from MovRef mr Inner Join Mov m on mr.nMovNro=m.nMovNro where nmovflag=0 and isnull(cagecodref,'')='' and m.copecod not like '42[12]12[234]')"
''SQL = SQL & " and m.copecod not like '7%'"
''SQL = SQL & " group by p.cPersCod,pei.cPersIDnro,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) , substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) , ddocfecha "
''SQL = SQL & " Union All "
''SQL = SQL & " select p.cPersCod,pei.cPersIDnro RUC,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) Serie, substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) Nro, ddocfecha,"
''SQL = SQL & " IsNull(sum(case when nmovimporte>0 then nmovimporte end),0) Bruto,"
''SQL = SQL & " IsNull(sum(case when nMovImporte < 0 And cctacontcod = '2114020201' then nmovimporte end),0) IMP "
''SQL = SQL & " from mov m"
''SQL = SQL & " Inner join MovCta mc on m.nmovnro=mc.nmovnro"
''SQL = SQL & " Inner Join OpeTpo o on m.copecod=o.copecod"
''SQL = SQL & " Inner join movdoc md on md.nmovnro=m.nmovnro and ndoctpo=2"
''SQL = SQL & " Inner Join MovGasto mg on mg.nmovnro=m.nMovNro"
''SQL = SQL & " Inner Join Persona p on p.cPersCod=mg.cPersCod"
''SQL = SQL & " Left Join PersID pei on pei.cPersCod=p.cPersCod and cPersIDTpo=2"
''SQL = SQL & " Where m.nmovflag = 0 And nmovestado = 10"
''SQL = SQL & " and cmovnro like '" & Format(MskAl.Text, "yyyymm") & "%'"
''SQL = SQL & " and m.nmovnro not in (select nmovnroref from MovRef mr Inner Join Mov m on mr.nMovNro=mr.nMovNro where nmovflag=0 and (cagecodref='' or cagecodref is null)  )"
''SQL = SQL & " and m.copecod like '70[12]105'"
''SQL = SQL & " group by p.cPersCod,m.nmovnro, pei.cPersIDnro,cpersNombre,left(cDocNro,PATINDEX('%-%',cDocNro) - 1) , substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) , ddocfecha order by cPersNombre"
'
''SQL = SQL & " Select pid.cperscod, pid.cPersIDNro RUC , p.cPersNombre, substring(AAA.cdocnro,1,3) SERIE   , substring(AAA.cdocnro,5, len(AAA.cdocnro)  )as NRO ,  AAA.dDocFecha ,   Total as Bruto , isnull(imp,0) IMP from     (Select b.cdocnro, a.nmovnro, sum(abs(nmovimporte)) Total  ,a.cMovDesc,CONVERT(varchar(10),b.dDocFecha,103) dDocFecha   from mov a  inner join movdoc b on a.nmovnro = b.nmovnro    inner join movcta c on c.nmovnro = b.nmovnro    "
''SQL = SQL & " where cmovnro like '" & CStr(Year(MskAl)) & Right("00" + CStr(Month(MskAl)), 2) & "' +'%' and nmovestado = 10 and nmovflag <> 1  and ndoctpo in ( 2,64) and (copecod like '7%'  or cOpeCod like '40%') and (cctacontcod like '25%' or cctacontcod like '45%' ) and  nMovImporte >0   group by b.cdocnro, a.nmovnro,a.cMovDesc,b.dDocFecha) As AAA Left Join      (Select a.nmovnro,  b.cdocnro, sum(abs(nmovimporte)) imp from mov a     inner join movdoc b on a.nmovnro = b.nmovnro    inner join movcta c on c.nmovnro = b.nmovnro"
''SQL = SQL & " where cmovnro like '" & CStr(Year(MskAl)) & Right("00" + CStr(Month(MskAl)), 2) & "'+'%' and nmovestado = 10 and nmovflag <> 1 and ndoctpo in ( 2,64) and (copecod like '7%'  or cOpeCod like '40%')  and cctacontcod = '2114020101'  group by b.cdocnro, a.nmovnro) As BBB     On AAA.nmovnro = BBB.nMovNro left join MovGasto mg ON mg.nMovNro=AAA.nMovNro left JOIN PErsona p ON p.cPersCod=mg.cPerscod  left JOIN PersID pid ON pid.cPersCod = p.cPerscod and pid.cPersIDTpo=2 Order by cPersNombre"
''dbo.GetNombreAPPaterno(cPersNombre) cApePat, dbo.GetNombreAPMaterno(cPersNombre) cApeMat, " & _
'        "replace(dbo.GetNombres(cPersNombre),' ','') cNombres, dbo.GetNombrePrimero(cPersNombre) cPrimerNombre,
'SQL = ""
''SQL = "SP_FinRentaCuarta_PDT '" & CStr(Year(MskAl)) & "', '" & CStr(Month(MskAl)) & "'"
''**Modificado por ALPA 20080214 ***********************
''SQL = "SP_FinRentaCuarta_PDT '" & CStr(Year(MskAl)) & "', '" & Format(CStr(Month(MskAl)), "00") & "'"
'SQL = "stp_sel_FinRentaCuarta_PDT '" & CStr(Year(MskAl)) & "', '" & Format(CStr(Month(MskAl)), "00") & "'"
''******************************************************
''Format(Right(cboMes.Text, 2), "00")
'Set rs = oConConsol.CargaRecordSet(SQL)
'Set rs.ActiveConnection = Nothing
'If rs.RecordCount > 0 Then
'    PB.Max = rs.RecordCount
'    PB.Visible = False
'End If
'
'If Not (rs.EOF And rs.BOF) Then
'    Open lsArc For Output As #NumeroArchivo
'    If LOF(1) > 0 Then
'        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
'            Close #1
'            Kill lsArc
'            Open lsArc For Append As #1
'        Else
'            Exit Sub
'        End If
'    End If
'    While Not rs.EOF
'        J = rs.RecordCount
'
'        If IsNull(rs!RUC) And rs!cPersNombre <> sNombres Then
'            rsTmp.AddNew
'            rsTmp!CodPers = rs!cPersCod
'        '    rsTmp!NomPers = rs!cApePat & Space(1) & rs!cApeMat & Space(1) & rs!cPrimerNombre
'        End If
'
'        sRUC = IIf(IsNull(rs!RUC), "", rs!RUC)
'        '/*Comentado por ALPA**/
'        '/*15/02/2008*/
'        'sApPaterno = PstaNombre(rs!cPersNombre, True)
'        '/***********/
'
'
'        'I = InStr(1, rs!cPersNombre, "/", vbTextCompare)
'        'If I = 0 Then
'        '    MsgBox "Corregir Tipo de Documento de: " & sApPaterno & Space(2) & sApMaterno, vbInformation, "Aviso"
'        '    PB.value = 0
'        '    PB.Visible = False
'        '    cmdGenerar.Enabled = False
'        '    Close #NumeroArchivo   ' Cierra el archivo.
'        '    oConConsol.CierraConexion
'        '    Set rsTmp = Nothing
'        '    Exit Sub
'        'Else
'        '    sApPaterno = Trim(Mid(rs!cPersNombre, 1, InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1))
'        'End If
'
'
'        'If InStr(1, rs!cPersNombre, "\", vbTextCompare) <> 0 Then
'        '    sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "\", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "\", vbTextCompare) - 1)
'        '
'        'Else
'        '    sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "/", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1)
'        'End If
'        'sNombres = LTrim(Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, ",", vbTextCompare) + 1, 100))
'
'        'sNroSerie = rs!Serie
'        'sNroRecibo = Right(rs!nro, 8)
'        'sFecEmision = Format(rs!dDocFecha, "dd/mm/yyyy")
'        'sMonto = Format(rs!Bruto, "####0.00")
'        'sRetencion = IIf(rs!IMP = 0, "0", "1")
'        'sPorcRet = "10"
'        'sRetIES = ""
'        'sNroOrden = ""
'        sTDOC = rs!TDOC
'        sNDOC = rs!NroDocId
'        '/*Comentado por ALPA**/
'        '/*15/02/2008*/
'        'sLinea = Trim(sRUC) & sP & Trim(sApPaterno) & sP & Trim(sApMaterno) & sP & Trim(sNombres) & sP & Trim(sNroSerie) & sP & Trim(sNroRecibo) & sP & Trim(sFecEmision) & sP & Trim(sMonto) & sP & Trim(sRetencion) & sP & Trim(sPorcRet) & sP & Trim(sRetIES) & sP & Trim(sNroOrden) & sP
'        sLinea = Trim(sTDOC) & sP & Trim(sNDOC) & sP & Trim(sRUC) & sP
'        '/*************/
'        Print #NumeroArchivo, sLinea
'
'        PB.value = PB.value + 1
'        rs.MoveNext
'    Wend
'    PB.value = 0
'    PB.Visible = False
'    cmdGenerar.Enabled = False
'    Close #NumeroArchivo   ' Cierra el archivo.
'    oConConsol.CierraConexion
'    MsgBox "Archivo Generado en su Spooler", vbInformation, "Aviso"
'    If RSVacio(rsTmp) Then Exit Sub
''    If rsTmp.RecordCount > 0 Then
''        MsgBox "El archivo presenta observaciones", vbInformation, "Aviso"
''        cmdImprimir.Enabled = True
''        cmdImprimir.SetFocus
''    End If
'Else
'    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
'    PB.Visible = False
'    If rsTmp.State = adStateOpen Then
'        rsTmp.Close
'    End If
'    Set rsTmp = Nothing
'
'End If
'
'Exit Sub
'ArchivoErr:
'    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
'    PB.Visible = False
'    cmdGenerar.Enabled = False
'    Close #NumeroArchivo   ' Cierra el archivo.
'    oConConsol.CierraConexion
'*******Fin Comentado
    
End Sub

Private Sub cmdImprimir_Click()
Dim cCodPers As String * 15
Dim cNomPers As String * 50
Dim sCodPersTmp As String
Dim oPrev As PrevioFinan.clsPrevioFinan
Set oPrev = PrevioFinan.clsPrevioFinan
Dim lsImpre As String

'Linea lsImpre, Cabecera(" P E R S O N A S   S I N   R U C   ", 0, gsSimbolo, gnColPage, "", Format(gdFecSis, gsFormatoFechaView)), , gnLinPage
Linea lsImpre, Cabecera(" P E R S O N A S   S I N   D O C U M E N T O  ", 0, gsSimbolo, gnColPage, "", Format(gdFecSis, gsFormatoFechaView)), , gnLinPage 'EJVG20121109
Linea lsImpre, String(75, "-"), , gnLinPage
Linea lsImpre, " CODIGO PERSONA                          NOMBRES Y APELLIDOS ", , gnLinPage
Linea lsImpre, String(75, "-"), , gnLinPage

If rsTmp.State = 0 Then
    MsgBox "Antes deberà haber procesado el archivo PDT.", vbInformation, "Aviso"
    Exit Sub
End If

If rsTmp.EOF Then
    MsgBox "No existen diferencias", vbInformation, "Aviso"
    Exit Sub
End If
rsTmp.MoveFirst
Do While Not rsTmp.EOF
    cCodPers = rsTmp!CodPers
    cNomPers = rsTmp!NomPers
    If Trim(sCodPersTmp) <> Trim(cCodPers) Then
        lsImpre = lsImpre & space(2) & cCodPers & space(10) & cNomPers & oImpresora.gPrnSaltoLinea
    End If
    sCodPersTmp = cCodPers
    rsTmp.MoveNext
Loop

rtfImp.Text = lsImpre
EnviaPrevio lsImpre, "REPORTE DE OBSERVACIONES", gnLinPage, True
End Sub

Private Sub cmdSalir_Click()

Unload Me
End Sub

Private Sub Form_Activate()
MskAl.SetFocus
End Sub

Private Sub Form_Load()
optDatos.Enabled = False
'optRuc.Enabled = False
CentraForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rsTmp.State = adStateOpen Then
    rsTmp.Close
End If
Set rsTmp = Nothing
End Sub

Private Sub MskAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValidaFecha(MskAl.Text) <> "" Then
        MsgBox "Fecha no válida...!", vbInformation, "Error"
        MskAl.SetFocus
    Else
        cmdGenerar.Enabled = True
        cmdGenerar.SetFocus
    End If
End If
End Sub


'''Option Explicit
'''Dim rsTmp As New ADODB.Recordset
'''
'''Private Sub cmdGenerar_Click()
'''Dim J As Integer
'''Dim lsArc As String
'''Dim sP As String * 1
'''Dim Sql As String
'''Dim NumeroArchivo As Integer
'''Dim sLinea As String
'''Dim rs As New ADODB.Recordset
'''Dim sRUC  As String * 11 '1
'''Dim sApPaterno  As String * 20 '2
'''Dim sApMaterno As String * 20 '3
'''Dim sNombres  As String * 20 '4
'''Dim sNroSerie As String * 3 '5
'''Dim sNroRecibo As String * 8 '6
'''Dim sFecEmision As String * 10 '7
'''Dim sMonto As String * 15 '8
'''Dim sRetencion As String * 1 '9
'''Dim sPorcRet As String  '10 ---- 10 %
'''Dim sRetIES As String * 1 '11
'''Dim sNroOrden As String * 1 '12
'''Dim sRUCTemp As String
'''Dim sNombreTemp As String
'''Dim oConConsol As DConecta
'''Set oConConsol = New DConecta
'''Dim i As Integer
'''
'''On Error GoTo ArchivoErr
'''
'''J = 0
'''sP = "|"
'''PB.Visible = True
'''PB.value = 0
'''PB.Min = 0
'''
'''If rsTmp.State = adStateOpen Then
'''    rsTmp.Close
'''End If
'''rsTmp.Fields.Append "CodPers", adChar, 15
'''rsTmp.Fields.Append "NomPers", adChar, 50
'''rsTmp.Open
''''AbreConexion
''''oConConsol.AbreConexionRemota "07", , , "03"
'''oConConsol.AbreConexion
'''
'''NumeroArchivo = FreeFile
'''lsArc = App.path & "\Spooler\0621" & gsRUC & Format(Me.MskAl, "YYYYMM") & ".txt"
'''
'''Sql = ""
'''Sql = "Select d.cPersCod, a.nmovnro,cPersNombre, dDocFecha,  " & _
'''        "(Select cPersIDnro from persid pid where pid.cPersCod = d.cperscod And cPersIDTpo = 2) RUC,  " & _
'''        "left(cDocNro,PATINDEX('%-%',cDocNro) - 1) Serie, substring(cDocNro,PATINDEX('%-%',cDocNro) + 1,50) Nro,  " & _
'''        "(Select sum(nMovImporte) from movcta mcta where mcta.nmovnro = a.nmovnro and nMovImporte > 0) Bruto,  " & _
'''        "(Select abs(nMovImporte) from movcta mcta where mcta.nmovnro = a.nmovnro and nMovImporte < 0 And cctacontcod = '2114020201' ) IMP  " & _
'''"from  mov a  " & _
'''"inner join movdoc b on a.nmovnro = b.nmovnro " & _
'''"inner join movgasto c on a.nmovnro = c.nmovnro  " & _
'''"inner join persona d on c.cperscod = d.cperscod  " & _
'''"where datediff(m, dDocFecha, '" & Format(MskAl.Text, "yyyy/mm/dd") & "') = 0  and ndoctpo = 2 and nmovflag = 0 and nMovEstado = 10 order by cPersNombre"
'''
''''dbo.GetNombreAPPaterno(cPersNombre) cApePat, dbo.GetNombreAPMaterno(cPersNombre) cApeMat, " & _
'''        "replace(dbo.GetNombres(cPersNombre),' ','') cNombres, dbo.GetNombrePrimero(cPersNombre) cPrimerNombre,
'''
'''Set rs = oConConsol.CargaRecordSet(Sql)
'''Set rs.ActiveConnection = Nothing
'''If rs.RecordCount > 0 Then
'''    PB.Max = rs.RecordCount
'''    PB.Visible = False
'''End If
'''
'''If Not (rs.EOF And rs.BOF) Then
'''    Open lsArc For Output As #NumeroArchivo
'''    If LOF(1) > 0 Then
'''        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
'''            Close #1
'''            Kill lsArc
'''            Open lsArc For Append As #1
'''        Else
'''            Exit Sub
'''        End If
'''    End If
'''    While Not rs.EOF
'''        J = rs.RecordCount
'''
'''        If IsNull(rs!RUC) And rs!cPersNombre <> sNombres Then
'''            rsTmp.AddNew
'''            rsTmp!CodPers = rs!cPersCod
'''        '    rsTmp!NomPers = rs!cApePat & Space(1) & rs!cApeMat & Space(1) & rs!cPrimerNombre
'''        End If
'''
'''        sRUC = IIf(IsNull(rs!RUC), "", rs!RUC)
'''
'''        sApPaterno = PstaNombre(rs!cPersNombre, True)
'''
'''
'''        i = InStr(1, rs!cPersNombre, "/", vbTextCompare)
'''        If i = 0 Then
'''            MsgBox "Corregir Tipo de Documento de: " & sApPaterno & Space(2) & sApMaterno, vbInformation, "Aviso"
'''            PB.value = 0
'''            PB.Visible = False
'''            cmdGenerar.Enabled = False
'''            Close #NumeroArchivo   ' Cierra el archivo.
'''            oConConsol.CierraConexion
'''            'Set rsTmp = Nothing
'''            Exit Sub
'''        Else
'''            sApPaterno = Trim(Mid(rs!cPersNombre, 1, InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1))
'''        End If
'''
'''
'''        If InStr(1, rs!cPersNombre, "\", vbTextCompare) <> 0 Then
'''            sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "\", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "\", vbTextCompare) - 1)
'''
'''        Else
'''            sApMaterno = Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, "/", vbTextCompare) + 1, InStr(1, rs!cPersNombre, ",", vbTextCompare) - InStr(1, rs!cPersNombre, "/", vbTextCompare) - 1)
'''        End If
'''        sNombres = LTrim(Mid(rs!cPersNombre, InStr(1, rs!cPersNombre, ",", vbTextCompare) + 1, 100))
'''
'''        sNroSerie = rs!Serie
'''        sNroRecibo = rs!nro
'''        sFecEmision = Format(rs!dDocFecha, "dd/mm/yyyy")
'''        sMonto = Format(rs!Bruto, "####0.00")
'''        sRetencion = IIf(IsNull(rs!IMP), "0", "1")
'''        sPorcRet = "10"
'''        sRetIES = ""
'''        sNroOrden = ""
'''
'''
'''        sLinea = Trim(sRUC) & sP & Trim(sApPaterno) & sP & Trim(sApMaterno) & sP & Trim(sNombres) & sP & Trim(sNroSerie) & sP & Trim(sNroRecibo) & sP & Trim(sFecEmision) & sP & Trim(sMonto) & sP & Trim(sRetencion) & sP & Trim(sPorcRet) & sP & Trim(sRetIES) & sP & Trim(sNroOrden) & sP
'''        Print #NumeroArchivo, sLinea
'''
'''        PB.value = PB.value + 1
'''        rs.MoveNext
'''    Wend
'''    PB.value = 0
'''    PB.Visible = False
'''    cmdGenerar.Enabled = False
'''    Close #NumeroArchivo   ' Cierra el archivo.
'''    oConConsol.CierraConexion
'''    MsgBox "Archivo Generado en su Spooler", vbInformation, "Aviso"
'''    If RSVacio(rsTmp) Then Exit Sub
''''    If rsTmp.RecordCount > 0 Then
''''        MsgBox "El archivo presenta observaciones", vbInformation, "Aviso"
''''        cmdImprimir.Enabled = True
''''        cmdImprimir.SetFocus
''''    End If
'''Else
'''    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
'''    PB.Visible = False
'''    If rsTmp.State = adStateOpen Then
'''        rsTmp.Close
'''    End If
'''    Set rsTmp = Nothing
'''    Exit Sub
'''
'''End If
'''
'''Exit Sub
'''ArchivoErr:
'''    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
'''    PB.Visible = False
'''    cmdGenerar.Enabled = False
'''    Close #NumeroArchivo   ' Cierra el archivo.
'''    oConConsol.CierraConexion
'''
'''
'''End Sub
'''
'''Private Sub cmdImprimir_Click()
'''Dim cCodPers As String * 15
'''Dim cNomPers As String * 50
'''Dim sCodPersTmp As String
'''Dim oPrev As Previo.clsPrevio
'''Set oPrev = Previo.clsPrevio
'''Dim lsImpre As String
'''
'''Linea lsImpre, Cabecera(" P E R S O N A S   S I N   R U C   ", 0, gsSimbolo, gnColPage, "", Format(gdFecSis, gsFormatoFechaView)), , gnLinPage
'''Linea lsImpre, String(75, "-"), , gnLinPage
'''Linea lsImpre, " CODIGO PERSONA                          NOMBRES Y APELLIDOS ", , gnLinPage
'''Linea lsImpre, String(75, "-"), , gnLinPage
'''If rsTmp.EOF Then
'''    MsgBox "No existen diferencias", vbInformation, "Aviso"
'''    Exit Sub
'''End If
'''rsTmp.MoveFirst
'''Do While Not rsTmp.EOF
'''    cCodPers = rsTmp!CodPers
'''    cNomPers = rsTmp!NomPers
'''    If Trim(sCodPersTmp) <> Trim(cCodPers) Then
'''        lsImpre = lsImpre & Space(2) & cCodPers & Space(10) & cNomPers & oImpresora.gPrnSaltoLinea
'''    End If
'''    sCodPersTmp = cCodPers
'''    rsTmp.MoveNext
'''Loop
'''
'''rtfImp.Text = lsImpre
'''EnviaPrevio lsImpre, "REPORTE DE OBSERVACIONES", gnLinPage, True
'''End Sub
'''
'''Private Sub cmdSalir_Click()
'''
'''Unload Me
'''End Sub
'''
'''Private Sub Form_Activate()
'''MskAl.SetFocus
'''End Sub
'''
'''Private Sub Form_Load()
'''CentraForm Me
'''
'''End Sub
'''
'''Private Sub Form_Unload(Cancel As Integer)
'''If rsTmp.State = adStateOpen Then
'''    rsTmp.Close
'''End If
'''Set rsTmp = Nothing
'''End Sub
'''
'''Private Sub MskAl_KeyPress(KeyAscii As Integer)
'''If KeyAscii = 13 Then
'''    If ValidaFecha(MskAl.Text) <> "" Then
'''        MsgBox "Fecha no válida...!", vbInformation, "Error"
'''        MskAl.SetFocus
'''    Else
'''        cmdGenerar.Enabled = True
'''        cmdGenerar.SetFocus
'''    End If
'''End If
'''End Sub
'EJVG20121109 ***
Private Sub darReporteCuarta_PrestadoresServicio()
    Dim rs As New ADODB.Recordset
    Dim sApPaterno  As String * 40
    Dim sApMaterno As String * 40
    Dim sNombres  As String * 40
    Dim sDomiciliado As String * 1
    Dim sConvDblTributo As String * 1
    Dim oConConsol As New DConecta
    Dim i As Integer, iMat As Long, indexMat As Long
    Dim MatPrestadorServicio() As Variant
    Dim bPrestadorServEncontrado As Boolean
    Dim lsCodPersona As String, lsNombres As String, lsTpoDoc As String, lsNroDoc As String
    
    On Error GoTo ArchivoErr
    
    j = 0
    sP = "|"
    PB.Visible = True
    PB.value = 0
    PB.Min = 0
    
    If rsTmp.State = adStateOpen Then
        rsTmp.Close
    End If
    rsTmp.Fields.Append "CodPers", adChar, 15
    rsTmp.Fields.Append "NomPers", adChar, 50
    rsTmp.Open
    oConConsol.AbreConexion
    NumeroArchivo = FreeFile
    
    lsArc = App.path & "\Spooler\0601" & Format(Me.MskAl, "YYYYMM") & gsRUC & ".ps4"
    sql = "stp_sel_FinRentaCuarta_4ta '" & CStr(Year(MskAl)) & "', '" & Format(CStr(Month(MskAl)), "00") & "'"

    Set rs = oConConsol.CargaRecordSet(sql)
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        PB.Max = rs.RecordCount
        PB.Visible = False
    End If
    
    iMat = 0
    ReDim MatPrestadorServicio(1 To 4, 0 To 0)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            bPrestadorServEncontrado = False
            For iMat = 1 To UBound(MatPrestadorServicio, 2)
                If MatPrestadorServicio(1, iMat) = rs!cPersCod Then
                    bPrestadorServEncontrado = True
                    Exit For
                End If
            Next
            If Not bPrestadorServEncontrado Then
                indexMat = UBound(MatPrestadorServicio, 2) + 1
                ReDim Preserve MatPrestadorServicio(1 To 4, 0 To indexMat)
                MatPrestadorServicio(1, indexMat) = rs!cPersCod
                MatPrestadorServicio(2, indexMat) = rs!cPersNombre
                MatPrestadorServicio(3, indexMat) = rs!TDOC
                MatPrestadorServicio(4, indexMat) = rs!NroDocId
            End If
            rs.MoveNext
        Loop
    End If
    
    If UBound(MatPrestadorServicio, 2) = 0 Then
        MsgBox "No existe Data para la exportacion", vbInformation, "Aviso"
        PB.Visible = False
        If rsTmp.State = adStateOpen Then
            rsTmp.Close
        End If
        Set rsTmp = Nothing
        Exit Sub
    End If
    
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    
    For iMat = 1 To UBound(MatPrestadorServicio, 2)
        lsCodPersona = MatPrestadorServicio(1, iMat)
        lsNombres = MatPrestadorServicio(2, iMat)
        lsTpoDoc = MatPrestadorServicio(3, iMat)
        lsNroDoc = MatPrestadorServicio(4, iMat)

        If lsNroDoc = "" And lsNombres <> sNombres Then
            rsTmp.AddNew
            rsTmp!CodPers = lsCodPersona
            rsTmp!NomPers = lsNombres
        End If

        sApPaterno = PstaNombre(lsNombres, True)

        i = InStr(1, lsNombres, "/", vbTextCompare)
        If i = 0 Then
            MsgBox "Corregir Tipo de Documento de: " & sApPaterno & space(2) & sApMaterno, vbInformation, "Aviso"
            PB.value = 0
            PB.Visible = False
            cmdGenerar.Enabled = False
            Close #NumeroArchivo   ' Cierra el archivo.
            oConConsol.CierraConexion
            Set rsTmp = Nothing
            Exit Sub
        Else
            sApPaterno = Trim(Mid(lsNombres, 1, InStr(1, lsNombres, "/", vbTextCompare) - 1))
        End If

        If InStr(1, lsNombres, "\", vbTextCompare) <> 0 Then
            sApMaterno = Mid(lsNombres, InStr(1, lsNombres, "\", vbTextCompare) + 1, InStr(1, lsNombres, ",", vbTextCompare) - InStr(1, lsNombres, "\", vbTextCompare) - 1)
        Else
            sApMaterno = Mid(lsNombres, InStr(1, lsNombres, "/", vbTextCompare) + 1, InStr(1, lsNombres, ",", vbTextCompare) - InStr(1, lsNombres, "/", vbTextCompare) - 1)
        End If

        sNombres = LTrim(Mid(lsNombres, InStr(1, lsNombres, ",", vbTextCompare) + 1, 100))
        sTDOC = lsTpoDoc
        sNDOC = lsNroDoc
        sDomiciliado = IIf(sTDOC = "01" Or sTDOC = "06", 1, 2)
        sConvDblTributo = 0
        
        
        sLinea = Trim(sTDOC) & sP & Trim(sNDOC) & sP & Trim(sApPaterno) & sP & Trim(sApMaterno) & sP & Trim(sNombres) & sP & Trim(sDomiciliado) & sP & Trim(sConvDblTributo) & sP
        
        Print #NumeroArchivo, sLinea

        PB.value = PB.value + 1
    Next
    PB.value = 0
    PB.Visible = False
    cmdGenerar.Enabled = False
    Close #NumeroArchivo
    oConConsol.CierraConexion
    MsgBox "Archivo Generado en su Spooler", vbInformation, "Aviso"
    Exit Sub
ArchivoErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    PB.Visible = False
    cmdGenerar.Enabled = False
    Close #NumeroArchivo   ' Cierra el archivo.
    oConConsol.CierraConexion
End Sub
'END EJVG *******
