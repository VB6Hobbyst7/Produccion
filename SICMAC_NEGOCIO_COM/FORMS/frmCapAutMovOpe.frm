VERSION 5.00
Begin VB.Form frmCapAutMovOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizacion para Movimientos Especiales"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
   Icon            =   "frmCapAutMovOpe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Persona"
      Height          =   705
      Left            =   5760
      TabIndex        =   10
      Top             =   4695
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton optTpo 
         Caption         =   "Agencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3045
         TabIndex        =   13
         Top             =   330
         Width           =   1140
      End
      Begin VB.OptionButton optTpo 
         Caption         =   "Areas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1770
         TabIndex        =   12
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton optTpo 
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   465
         TabIndex        =   11
         Top             =   330
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   12960
      TabIndex        =   5
      Top             =   5625
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   14190
      TabIndex        =   4
      Top             =   5625
      Width           =   975
   End
   Begin SICMACT.FlexEdit FlexData 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7646
      Cols0           =   12
      HighLight       =   1
      EncabezadosNombres=   "#-Operacion-Cuenta-Persona-Monto Solic.-Monto Aprob.-Moneda-Observaciones-Estado-Id-MontoFinSol-MontoFinDol"
      EncabezadosAnchos=   "300-4000-2500-5700-1200-1200-800-3000-0-1400-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-5-X-7-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      BackColor       =   16777215
      EncabezadosAlineacion=   "C-L-L-L-R-R-L-L-L-R-R-C"
      FormatosEdit    =   "0-0-0-0-2-2-0-0-1-3-2-2"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483625
      CellBackColor   =   16777215
   End
   Begin VB.Frame FraData 
      Height          =   5385
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   15255
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Refrescar"
         Height          =   375
         Left            =   4335
         TabIndex        =   9
         Top             =   4860
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   375
         Left            =   2970
         TabIndex        =   8
         Top             =   4860
         Width           =   855
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Rechazar"
         Height          =   375
         Left            =   1605
         TabIndex        =   3
         Top             =   4860
         Width           =   855
      End
      Begin VB.CommandButton cmdAprobar 
         Caption         =   "Aprobar"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   4860
         Width           =   855
      End
      Begin VB.Label lblEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   11445
         TabIndex        =   7
         Top             =   4800
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10680
         TabIndex        =   6
         Top             =   4800
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmCapAutMovOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Const OPEN_EXISTING = 3
'Private Const GENERIC_WRITE = &H40000000
'Private Const FILE_SHARE_READ = &H1
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const INVALID_HANDLE_VALUE = -1
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
'Public Function EnviarMensajeA(ByVal Emisor As String, ByVal Receptor As String, ByVal Mensaje As String)
'Dim lngH As Long, strTextoAEnviar As String, lngResult As Long
'   strTextoAEnviar = Emisor & Chr(0) & Receptor & Chr(0) & Mensaje & Chr(0)
'   lngH = CreateFile("\\" & Receptor & "\mailslot\messngr", GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
'   If lngH = INVALID_HANDLE_VALUE Then
'      EnviarMensaje = 0
'   Else
'      If WriteFile(lngH, strTextoAEnviar, Len(strTextoAEnviar), lngResult, 0) = 0 Then
'         EnviarMensaje = 0
'      Else
'         EnviarMensaje = lngResult
'      End If
'      CloseHandle lngH
'   End If
'End Function
'Private Sub CmdButton1_Click()
'  Call EnviaMensaje("200300", 5000, 1)
'End Sub
'Public Sub EnviaMensaje(ByVal cOpeCod, ByVal nMonto As Double, ByVal nMoneda As Integer)
'Dim ssql As String, saux As String
'Dim rs As ADODB.Recordset, oconecta As DConecta
'  Set oconecta = New DConecta
'  Set rs = New ADODB.Recordset
'    If nMoneda = gMonedaNacional Then
'        saux = " and (c.nmontofinsol)>=" & nMonto
'    ElseIf nMoneda = gMonedaExtranjera Then
'        saux = " and (c.nmontofinsol)>=" & nMonto
'    End If
'
'   ssql = "Select rh.cuser,cl.cworkstation from  rrhh  rh "
'   ssql = ssql & " inner join (Select Max(cast(convert(char(8),drhcargoFecha,112) as int)) as Fecha,cperscod,crhcargocod from rhcargos "
'   ssql = ssql & " group by cperscod,crhcargocod) rhc on rhc.cperscod=rh.cperscod "
'   ssql = ssql & " inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
'   ssql = ssql & " inner join caplogeo cl on cl.cuser=rh.cuser "
'   ssql = ssql & " where  rh.nrhestado='201' and c.copecod='" & cOpeCod & "'" & saux
'
'    Set oconecta = New DConecta
'    oconecta.AbreConexion
'    Set rs = oconecta.CargaRecordSet(ssql)
'    oconecta.CierraConexion
'    Set oconecta = Nothing
'    If rs.State = 1 Then
'        While Not rs.EOF
'          ' Shell "Net send " & rs!cworkstation & " URGENTE: Tiene Operacion Pendiente !!!"
'          EnviarMensajeA Vusuario, rs!cworkstation, "URGENTE: Tiene Operaciones Pendientes !!!"
'            rs.MoveNext
'        Wend
'    End If
'Exit Sub
'MensaError:
'    Call RaiseError(MyUnhandledError, "frmCapAutorizacionPrueba: CargaOperaciones  Method")
'End Sub


'*******************************************************************************************
'Estados de MovCapAutorizacion:
'P-->Pendiente,A-->Aprobada,E-->Efectuada,R-->Rechazada
'*******************************************************************************************
Option Explicit
Dim VEditob As Boolean, VSalirb As Boolean

Private Sub Carga_Flex()
  Dim sSql As String, rs As ADODB.Recordset, i As Long
  Dim oconecta As DConecta
  
  Set rs = New ADODB.Recordset
  
  On Error GoTo MensaError
  
  'where (ch.copecod like  '[23][0-2][0123]' or ch.copecod like '90002[0-6]' or ch.copecod like  '9010[01][0-9]' or ch.copecod like '90003[0-5]' )
  
'  If optTpo(0).value Then
            sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,p.cpersnombre + space(150-len(p.cpersnombre))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
            sSql = sSql & " from capautorizacionope m "
            sSql = sSql & "  inner join (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion "
            sSql = sSql & " from opetpo ch "
            sSql = sSql & " inner join opetpo f on f.copecod=case when left(ch.copecod,1)='2' then  left(ch.copecod,2)+'0000' when  left(ch.copecod,2)='30' then  left(ch.copecod,2)+'0000'  when left(ch.copecod,2)='31' then  left(ch.copecod,2)+'0000'      when  left(ch.copecod,5)='90002' then '900020'  when left(ch.copecod,5)='90100' then  '901001' when left(ch.copecod,5)='90101' then  '901011' when  left(ch.copecod,5)='90003'  then   '900030'   end "
            sSql = sSql & " where ch.copecod like  '2%[1-9]00' or ch.copecod like  '3[0-1]%00' or ch.copecod like '90002[1-6]' or ch.copecod like  '9010[01][2-9]' or ch.copecod like '90003[1-5]' ) o on o.copecod=m.copecodori "
            sSql = sSql & " inner join persona p on p.cperscod=m.cperscodcli "
            sSql = sSql & "  inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
            sSql = sSql & "  inner join capautorizacionrango c on c.copecod=m.copecodori "
            sSql = sSql & "  inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & Vusuario & "' and  rh1.nrhestado='201' "
            sSql = sSql & "  order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
            sSql = sSql & "  where left(m.cultimaactualizacion,8)=convert(char(8),getdate(),112) and m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
            sSql = sSql & " order by m.nidaut  "
            
'  ElseIf optTpo(1).value Then
'            sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,substring(m.cperscodcli,4,2) + ao.careadescripcion + '-' + ad.careadescripcion + space(150-len(substring(m.cperscodcli,4,2) + ao.careadescripcion + '-' + ad.careadescripcion))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
'            sSql = sSql & "  from capautorizacionope m "
'            sSql = sSql & " inner join (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion "
'            sSql = sSql & " from opetpo ch "
'            sSql = sSql & " inner join opetpo f on f.copecod=case when left(ch.copecod,1)='2' then  left(ch.copecod,2)+'0000' when  left(ch.copecod,2)='30' then  left(ch.copecod,2)+'0000'  when left(ch.copecod,2)='31' then  left(ch.copecod,2)+'0000'      when  left(ch.copecod,5)='90002' then '900020'  when left(ch.copecod,5)='90100' then  '901001' when left(ch.copecod,5)='90101' then  '901011' when  left(ch.copecod,5)='90003'  then   '900030'   end "
'            sSql = sSql & " where ch.copecod like  '2%[1-9]00' or ch.copecod like  '3[0-1]%00' or ch.copecod like '90002[1-6]' or ch.copecod like  '9010[01][2-9]' or ch.copecod like '90003[1-5]' ) o on o.copecod=m.copecodori "
'            sSql = sSql & " inner join areas ao on ao.careacod=left(m.cperscodcli,3) "
'            sSql = sSql & " inner join areas ad on ad.careacod=right(m.cperscodcli,3) "
'            sSql = sSql & " inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
'            sSql = sSql & " inner join capautorizacionrango c on c.copecod=m.copecodori "
'            sSql = sSql & " inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & Vusuario & "' and  rh1.nrhestado='201' "
'            sSql = sSql & " order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
'            sSql = sSql & " where left(m.cultimaactualizacion,8)=convert(char(8),getdate(),112) and m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
'            sSql = sSql & " order by m.nidaut "
'
'
'  ElseIf optTpo(2).value Then
'
'            sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,ao.cagedescripcion + '-' + ad.cagedescripcion + space(150-len( ao.cagedescripcion + '-' + ad.cagedescripcion))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
'            sSql = sSql & " from capautorizacionope m "
'            sSql = sSql & " inner join (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion "
'            sSql = sSql & " from opetpo ch "
'            sSql = sSql & " inner join opetpo f on f.copecod=case when left(ch.copecod,1)='2' then  left(ch.copecod,2)+'0000' when  left(ch.copecod,2)='30' then  left(ch.copecod,2)+'0000'  when left(ch.copecod,2)='31' then  left(ch.copecod,2)+'0000'      when  left(ch.copecod,5)='90002' then '900020'  when left(ch.copecod,5)='90100' then  '901001' when left(ch.copecod,5)='90101' then  '901011' when  left(ch.copecod,5)='90003'  then   '900030'   end "
'            sSql = sSql & " where ch.copecod like  '2%[1-9]00' or ch.copecod like  '3[0-1]%00' or ch.copecod like '90002[1-6]' or ch.copecod like  '9010[01][2-9]' or ch.copecod like '90003[1-5]' ) o on o.copecod=m.copecodori "
'            sSql = sSql & " inner join agencias ao on ao.cagecod=left(m.cperscodcli,2) "
'            sSql = sSql & " inner join agencias ad on ad.cagecod=right(m.cperscodcli,2)"
'            sSql = sSql & " inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
'            sSql = sSql & " inner join capautorizacionrango c on c.copecod=m.copecodori "
'            sSql = sSql & " inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & Vusuario & "' and  rh1.nrhestado='201' "
'            sSql = sSql & " order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
'            sSql = sSql & " where left(m.cultimaactualizacion,8)=convert(char(8),getdate(),112) and m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
'            sSql = sSql & " order by m.nidaut "
'
'
'
'  End If
   
'   sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,p.cpersnombre + space(150-len(p.cpersnombre))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
'   sSql = sSql & " from capautorizacionope m "
'   sSql = sSql & "  inner join (Select f.copedesc +':'+ ch.copedesc as operacion, ch.copecod  from opetpo ch  inner join opetpo f on f.copecod=left(ch.copecod,2)+'0000'  where ch.copecod like  '2%[1-9]00'  ) o on o.copecod=m.copecodori "
'   sSql = sSql & " inner join persona p on p.cperscod=m.cperscodcli "
'   sSql = sSql & "  inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
'   sSql = sSql & "  inner join capautorizacionrango c on c.copecod=m.copecodori "
'   sSql = sSql & "  inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & Vusuario & "' and  rh1.nrhestado='201' "
'   sSql = sSql & "  order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
'   sSql = sSql & "  where left(m.cultimaactualizacion,8)=convert(char(8),getdate(),112) and m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
'   sSql = sSql & " order by m.nidaut  desc"

'  ssql = "exec Cap_MuestaMovOpePendientes '" & Vusuario & "'"
  ' Opera,Cuenta, persona,m.nmontosolicitado,m.nmontoaprobado,moneda,m.cautobs,m.cautestado,m.nidaut
  
 
    Set oconecta = New DConecta
    oconecta.AbreConexion
    rs.CursorLocation = adUseClient
    Set rs = oconecta.CargaRecordSet(sSql)
    oconecta.CierraConexion
    Set oconecta = Nothing

    FlexData.Clear
    
   
   i = 1
    If rs.State = 1 Then
       If rs.RecordCount > 0 Then
               FlexData.rsFlex = rs
               Do Until i > (FlexData.Rows - 1)
                   If Right(FlexData.TextMatrix(i, 6), 1) = gMonedaNacional Then
                      FlexData.Row = i
                      Call FlexData.BackColorRow(&HC0FFFF)
            
                   Else
                      FlexData.Row = i
                      Call FlexData.BackColorRow(&HC0FFC0)
                   End If
                   FlexData.TextMatrix(i, 4) = Format(FlexData.TextMatrix(i, 4), "#,###,##0.00")
                   FlexData.TextMatrix(i, 5) = Format(FlexData.TextMatrix(i, 5), "#,###,##0.00")
                    i = i + 1
            
                Loop
            
                If FlexData.Rows - 1 > 0 Then
                    FlexData.Row = 1
                    lblEstado.Caption = IIf(FlexData.TextMatrix(FlexData.Row, 8) = "P", "Pendiente", IIf(FlexData.TextMatrix(FlexData.Row, 8) = "A", "Aprobada", "Rechazada"))
                End If
              If rs.State = 1 Then rs.Close
        Else
                Cabecera
                FlexData.Rows = 2
        End If
    Else
      Cabecera
    End If
    Set rs = Nothing
  Exit Sub

MensaError:
    Call RaiseError(MyUnhandledError, "frmCapAutorizacionPrueba: CargaOperaciones  Method")
      
End Sub
Private Sub Cabecera()
            FlexData.TextMatrix(0, 0) = "#"
            FlexData.TextMatrix(0, 1) = "Operacion"
            FlexData.TextMatrix(0, 2) = "Cuenta"
            FlexData.TextMatrix(0, 3) = "Persona"
            FlexData.TextMatrix(0, 4) = "Monto Solic."
            FlexData.TextMatrix(0, 5) = "Monto Aprob."
            FlexData.TextMatrix(0, 6) = "Moneda"
            FlexData.TextMatrix(0, 7) = "Observaciones"
            FlexData.TextMatrix(0, 8) = "Estado"
            FlexData.TextMatrix(0, 9) = "Id"
            FlexData.TextMatrix(0, 10) = "MontoFinSol"
            FlexData.TextMatrix(0, 11) = "MontoFinDol"

End Sub

Private Sub cmdAprobar_Click()
If FlexData.Rows - 1 > 0 Then
  If Val(FlexData.TextMatrix(FlexData.Row, 5)) = 0 Then
   MsgBox "Si desea aprobar este movimiento el monto aprobado debe ser mayor que cero", vbOKOnly + vbInformation, "Atención"
   Exit Sub
  End If
    FlexData.TextMatrix(FlexData.Row, 8) = "A"
    lblEstado.Caption = IIf(FlexData.TextMatrix(FlexData.Row, 8) = "P", "Pendiente", IIf(FlexData.TextMatrix(FlexData.Row, 8) = "A", "Aprobada", "Rechazada"))
    cmdCancelar.Enabled = True
    VEditob = True
End If
End Sub

Private Sub cmdCancelar_Click()
If MsgBox("¿Desea cancelar las modificaciones hechas en este formulario?", vbYesNo + vbQuestion, "Atención") = vbYes Then
   Carga_Flex
   cmdCancelar.Enabled = False
End If
End Sub

Private Sub cmdGrabar_Click()
 Dim aut As DAutorizacion, rs As ADODB.Recordset, herror As Boolean, VUltAct As String
 Dim cOpeCod As String, cCtaCod As String, cPersCod As String, nMontoSolicitado As Double
 Dim nMontoAprobado As Double, nMoneda As Integer, cAutObs As String, cautestado As String, nIdaut As Long
 Dim i As Long
       
    Set aut = New DAutorizacion
    
    herror = False
     
        
  Dim cnromov As New DMov
  Dim Valor As String
  Valor = cnromov.GeneraMovNro(Format(Date, "yyyy-mm-dd hh:mm:ss"), "01", Vusuario)
 
    i = 1
        
        'Opera,Cuenta,persona,nmontosolicitado,nmontoaprobado,moneda, cautobs, cautestado nidaut
        
        Do While i <= (FlexData.Rows - 1)
         With FlexData
                  
           cOpeCod = Right(.TextMatrix(i, 1), 6)
           cCtaCod = Trim(.TextMatrix(i, 2))
           cPersCod = Right(.TextMatrix(i, 3), 13)
           nMontoSolicitado = CDbl(IIf(.TextMatrix(i, 4) = "", "0.00", .TextMatrix(i, 4)))
           nMontoAprobado = CDbl(IIf(.TextMatrix(i, 5) = "", "0.00", .TextMatrix(i, 5)))
           nMoneda = CInt(Right(.TextMatrix(i, 6), 1))
           cAutObs = CStr(Trim(.TextMatrix(i, 7)))
           cautestado = CStr(Trim(.TextMatrix(i, 8)))
           nIdaut = CLng(.TextMatrix(i, 9))
           
           
           VUltAct = Valor
           VUltAct = VUltAct  'IIf(CStr(.TextMatrix(i, 10)) = False, "0", "1")
           
         End With
            
                 herror = IIf(aut.AMovAutorizacionApro(nIdaut, cCtaCod, cPersCod, cOpeCod, "280715", nMontoSolicitado, nMontoAprobado, nMoneda, cautestado, cAutObs, "", Date, "MPBR", Date, "", VUltAct), herror, True)
                 
                'herror = IIf(AMovAutorizacionApro(nIdAut, cCtaCod, cPersCod, copecod, "280715", nMontoSolicitado, nMontoAprobado, nMoneda, cautestado, cAutObs, "", Date, "MPBR", Date, "", VUltAct) = False, herror, True)
          
            i = i + 1
        Loop
        
    Carga_Flex
    VEditob = False
    Set aut = Nothing
      
End Sub


Public Function AMovAutorizacionApro(ByVal nIdaut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal cOpeCod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovNro As String, cUltimaActualizacion As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
  Dim oconecta As DConecta
      AMovAutorizacionApro = False
      
      On Error GoTo MensaErr
        sSql = "exec Cap_ManMovAutorizacion_sp " & nIdaut & ",'" & cCtaCod & "','" & cPersCodCli & "','" & cOpeCod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovNro & "','" & cUltimaActualizacion & "'"
      
        
           Set oconecta = New DConecta
           Set rs = New ADODB.Recordset
           oconecta.AbreConexion
           Set rs = oconecta.Ejecutar(sSql)
           If rs.State = 1 Then
                If Not (rs.EOF Or rs.BOF) Then
                    If rs.Fields(0).value = nIdaut Then AMovAutorizacionApro = True
                End If
                 rs.Close
           End If
           
           Set rs = Nothing
           oconecta.CierraConexion
           Set oconecta = Nothing
      
      Exit Function
      
MensaErr:
      Call RaiseError(MyUnhandledError, "DAutorizacion:AMovAutorizacionApro Method")
End Function
Private Sub cmdRechazar_Click()
If FlexData.Rows - 1 > 0 Then
    If Val(FlexData.TextMatrix(FlexData.Row, 5)) > 0 Then
        MsgBox "Si desea rechazar este movimiento el monto aprobado debe ser igual a cero", vbOKOnly + vbInformation, "Atención"
        Exit Sub
    End If
    FlexData.TextMatrix(FlexData.Row, 8) = "R"
    lblEstado.Caption = IIf(FlexData.TextMatrix(FlexData.Row, 8) = "P", "Pendiente", IIf(FlexData.TextMatrix(FlexData.Row, 8) = "A", "Aprobada", "Rechazada"))
    VEditob = True
    cmdCancelar.Enabled = True
End If

End Sub

Private Sub cmdRefrescar_Click()
If MsgBox("¿Desea refrescar la informacion de este formulario?", vbYesNo + vbQuestion, "Atención") = vbYes Then
   Carga_Flex
   cmdCancelar.Enabled = False
   VEditob = False
End If

End Sub

Private Sub cmdsalir_Click()
VSalirb = True

    If VEditob = True Then
        Call cmdGrabar_Click
    End If
    Unload Me

End Sub



Private Sub FlexData_Click()
  If FlexData.Rows - 1 > 0 Then
    lblEstado.Caption = IIf(FlexData.TextMatrix(FlexData.Row, 8) = "P", "Pendiente", IIf(FlexData.TextMatrix(FlexData.Row, 8) = "A", "Aprobada", "Rechazada"))
  End If
End Sub

Private Sub FlexData_OnCellChange(pnRow As Long, pnCol As Long)
VEditob = True
 Select Case pnCol
   Case 4
       'FlexData.TextMatrix(pnRow, pnCol) = Format(FlexData.TextMatrix(pnRow, pnCol), "#,###,##0.00")
   Case 5
       FlexData.TextMatrix(pnRow, pnCol) = Format(FlexData.TextMatrix(pnRow, pnCol), "#,###,##0.00")
       If CDbl((FlexData.TextMatrix(pnRow, 5)) > CDbl(FlexData.TextMatrix(pnRow, 4))) And FlexData.TextMatrix(pnRow, 4) <> "" And FlexData.TextMatrix(pnRow, 4) <> "0.00" Then
                MsgBox "El Monto Aprobado debe ser menor o igual que el monto solicitado", vbInformation, "Atención"
                FlexData.TextMatrix(pnRow, pnCol) = "0.00"
       End If
       
 End Select
End Sub

Private Sub FlexData_OnRowChange(pnRow As Long, pnCol As Long)
 If FlexData.Rows - 1 > 0 Then
    lblEstado.Caption = IIf(FlexData.TextMatrix(FlexData.Row, 8) = "P", "Pendiente", IIf(FlexData.TextMatrix(FlexData.Row, 8) = "A", "Aprobada", "Rechazada"))
  End If
End Sub

Private Sub Form_Load()
 Me.Top = 0
 Me.Left = 0
 VEditob = False
 VSalirb = False
 Carga_Flex
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VSalirb = False Then
        If VEditob = True Then
            Call cmdGrabar_Click
        End If
        Unload Me
    
 End If
 
End Sub

'Private Sub optTpo_Click(Index As Integer)
'    If VEditob = True Then
'           Call cmdGrabar_Click
'     End If
'     Carga_Flex
'
'End Sub
