VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogCalculaAsientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regeneración de Asientos contables - Almacén"
   ClientHeight    =   4695
   ClientLeft      =   1965
   ClientTop       =   2460
   ClientWidth     =   7485
   Icon            =   "frmLogCalculaAsientos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   4080
      Width           =   1155
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   4725
      TabIndex        =   1
      Top             =   4080
      Width           =   1155
   End
   Begin VB.Frame fraCalcula 
      Height          =   4515
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   7290
      Begin VB.Frame fraMes 
         Height          =   1395
         Left            =   1260
         TabIndex        =   8
         Top             =   2340
         Width           =   4815
         Begin VB.TextBox txtAnio 
            Height          =   315
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   10
            Top             =   300
            Width           =   855
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   300
            Width           =   2295
         End
         Begin Sicmact.TxtBuscar txtAlmacen 
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   840
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            Appearance      =   0
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
            sTitulo         =   ""
         End
         Begin VB.Label lblAlmacenG 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Top             =   840
            Width           =   2490
         End
         Begin VB.Label lblAlmacen 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mes y Año"
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
            Left            =   300
            TabIndex        =   11
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fraBarra 
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   795
         Left            =   1260
         TabIndex        =   12
         Top             =   2340
         Visible         =   0   'False
         Width           =   4815
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "En el caso de los bienes que si se controlan por serie, se asume el valor con el que ingresaron, pero prorrateado bien por bien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   780
         TabIndex        =   7
         Top             =   1680
         Width           =   6180
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vuelve a calcular los valores promedio de los bienes que no tienen control por serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   780
         TabIndex        =   6
         Top             =   1080
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vuelve a generar los asientos contables de las Salidas de Almacén"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   780
         TabIndex        =   5
         Top             =   660
         Width           =   5550
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Este proceso realiza lo siguiente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Este proceso puede durar varios minutos.                              Al finalizar se muestran los asientos contables en el previo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3900
         Width           =   4305
      End
   End
End
Attribute VB_Name = "frmLogCalculaAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAlmacen As ADODB.Recordset
Dim rsProductos As ADODB.Recordset
Dim lsTitulo As Boolean
Dim lsCadena As String

Private Sub Form_Load()
CentraForm Me
txtAnio = Year(Date)
cboMes.AddItem "Seleccione mes -----"
cboMes.AddItem "ENERO"
cboMes.AddItem "FEBRERO"
cboMes.AddItem "MARZO"
cboMes.AddItem "ABRIL"
cboMes.AddItem "MAYO"
cboMes.AddItem "JUNIO"
cboMes.AddItem "JULIO"
cboMes.AddItem "AGOSTO"
cboMes.AddItem "SEPTIEMBRE"
cboMes.AddItem "OCTUBRE"
cboMes.AddItem "NOVIEMBRE"
cboMes.AddItem "DICIEMBRE"
cboMes.ListIndex = 0

Dim oDoc As DOperaciones
Set oDoc = New DOperaciones
Me.txtAlmacen.rs = oDoc.GetAlmacenes
cboMes.ListIndex = 0
End Sub

Private Sub txtAlmacen_EmiteDatos()
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
End Sub

Private Sub cmdProcesar_Click()
     Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsBS As ADODB.Recordset
    Set rsBS = New ADODB.Recordset
    Dim lnSaldoIni As Double
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oMov As DMov
    Set oMov = New DMov
    Dim lsCtaCnt As String
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim lsAsientos As String
    Dim lnItem As Long
    Dim lsCaption As String
    Dim lsCtaTemp As String
    Dim lnAgenciaCod As Integer
    
    Dim nMes As Integer, nAnio As Integer
    Dim cFecIni As String, cFecFin As String
    
    lsAsientos = ""
    lsCadena = Caption
    
    If cboMes.ListIndex <= 0 Then
       MsgBox "Debe indicar un mes válido..." + Space(10), vbInformation, "Aviso"
       Exit Sub
    End If
    
    If txtAlmacen.Text = "" Then
       MsgBox "Debe indicar una Agencia válido..." + Space(10), vbInformation, "Aviso"
       Exit Sub
    End If
    
    'If Len(txtAnio) = 0 Or Val(txtAnio) < Year(Date) Then
    If Len(txtAnio) = 0 Then
       MsgBox "Debe indicar un año válido..." + Space(10), vbInformation, "Aviso"
       Exit Sub
    End If
    
    nMes = cboMes.ListIndex
    nAnio = CInt(txtAnio)
    lnAgenciaCod = txtAlmacen.Text
    
    cFecIni = DateSerial(nAnio, nMes + 0, 1)
    cFecFin = DateSerial(nAnio, nMes + 1, 0)
    
    If YaHaySaldos(nMes, nAnio, lnAgenciaCod) Then
       MsgBox "Ya se ha procesado " & cboMes.Text & " - " & txtAnio & "..." + Space(10), vbInformation, "Aviso"
       Exit Sub
    End If
    
'    If Not IsDate(Me.mskFecIni.Text) Then
'        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
'        mskFecIni.SetFocus
'        Exit Sub
'    ElseIf Not IsDate(Me.mskFecFin.Text) Then
'        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
'        mskFecFin.SetFocus
'        Exit Sub
'    End If
    
    If MsgBox("¿ Está seguro de volver a generar los Asientos Contables de las" & Space(10) & vbCrLf & Space(15) + "Salidas de Almacen (Todos los Almacenes) para " & cboMes.Text & " - " & txtAnio & " ? " & vbCrLf & vbCrLf & Space(15) + "Este proceso puede durar varios minutos", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsCaption = Caption
    fraMes.Visible = False
    fraBarra.Visible = True
    cmdProcesar.Visible = False
    CmdSalir.Visible = False
    oCon.AbreConexion
    
    'Procesar
    sql = " Select Distinct MD.cDocNro, md.dDocFecha, M.cOpeCod, M.nMovNro, M.cMovNro, MO.cAreaCod, MO.cAgeCod From MovDoc MD " _
        & " Inner Join Mov M On MD.nMovNro =  M.nMovNro" _
        & " Left  Join MovObjAreaAgencia MO On MO.nMovNro =  M.nMovNro" _
        & " Where nDocTpo = '71' And M.nMovNro In (Select Max(MMD.nMovNro) From MovDoc MMD Where MMD.cDocNro = MD.cDocNro And nDocTpo = '71')" _
        & " And dDocFecha Between '" & Format(CDate(cFecIni), gsFormatoFecha) & "' And '" & Format(CDate(cFecFin), gsFormatoFecha) & "' And M.nMovEstado = 10 And M.nMovFlag Not In (1,2,3,5) and MO.cAgeCod= " & lnAgenciaCod & " " _
        & " " _
        & " Order by m.cMovNro"
    
    'Procesar
    'sql = " Select Distinct MD.cDocNro, md.dDocFecha, M.cOpeCod, M.nMovNro, M.cMovNro, MO.cAreaCod, MO.cAgeCod From MovDoc MD " _
        & " Inner Join Mov M On MD.nMovNro =  M.nMovNro" _
        & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
        & " Left  Join MovObjAreaAgencia MO On MO.nMovNro =  M.nMovNro" _
        & " Where nDocTpo = '71' And M.nMovNro In (Select Max(MMD.nMovNro) From MovDoc MMD Where MMD.cDocNro = MD.cDocNro And nDocTpo = '71')" _
        & " And dDocFecha Between '" & Format(CDate(cFecIni), gsFormatoFecha) & "' And '" & Format(CDate(cFecFin), gsFormatoFecha) & "' And M.nMovEstado = 10 And M.nMovFlag Not In (1,2,3,5) And" _
        & " MBS.cBSCod In ('1110100701') " _
        & " Order by md.dDocFecha"
    
    Set rs = oCon.CargaRecordSet(sql)
     
    Me.PrgBar.value = 0
    Me.PrgBar.Max = rs.RecordCount + 2
    
   ' oCon.BeginTrans
    
    While Not rs.EOF
        sql = " Delete MovCta Where nMovNro = " & rs!nMovNro
        oCon.Ejecutar sql
        
        'sql = " Select cBSCod, MBS.nMovItem , nMovCant, dbo.LogGetPrecioPromedioFecha(cBSCod,'" & rs!dDocFecha & "',(Select Top 1 nProdTpo from LogOpeProd Where cOpeCod = '" & rs!cOpeCod & "')) Monto, cCtaContCod From MovBS MBS " _
            & " Inner Join MovCant MC On MBS.nMovNro = MC.nMovNro And MBS.nMovItem = MC.nMovItem" _
            & " Inner Join CTABS CTABS On cBSCod LIKE COBJETOCOD + '%' And CTABS.cOpeCod = '" & rs!cOpeCod & "'" _
            & " Where MBS.nMovNro = " & rs!nMovNro & ""
        
        sql = " Select cBSCod, MBS.nMovItem , nMovCant, dbo.LogGetPrecioPromedioFecha(cBSCod,'" & Format(rs!dDocFecha, gsFormatoFecha) & "',(Select Top 1 nProdTpo from LogOpeProd Where cOpeCod = '" & rs!COPECOD & "'),'" & rs!cMovNro & "',nMovCant, MBS.nMovBSOrden) Monto, (Select Top 1 Replace(cCtaContCod,'AG',  Right('0' + Ltrim(str(MBS.nMovBSOrden)),2) ) cCtaContCod From CTABS CTABS Where MBS.cBSCod LIKE COBJETOCOD + '%' And cOpeCod = '" & rs!COPECOD & "' order by len(COBJETOCOD) desc) cCtaContCod,  dbo.GetLogMontoSalidaAFBND(MBS.cBSCod, MBS.nMovNro, MBS.nMovItem) P_AFBND, MBS.nMovBSOrden nAlmacenCod From MovBS MBS " _
            & " Inner Join MovCant MC On MBS.nMovNro = MC.nMovNro And MBS.nMovItem = MC.nMovItem" _
            & " " _
            & " Where MBS.nMovNro = " & rs!nMovNro & ""
        Set rsBS = oCon.CargaRecordSet(sql)
        
        Caption = lsCaption & " /  " & rs!nMovNro & " - " & rs!dDocFecha
        
        If rsBS.EOF And rsBS.BOF Then
            MsgBox "no existe cuenta contable"
        Else
            While Not rsBS.EOF
                If Left(rsBS!cBSCod, 3) = "112" Or Left(rsBS!cBSCod, 3) = "113" Then
                    If rsBS!P_AFBND <> 0 Then
                        oMov.InsertaMovCta rs!nMovNro, rsBS!nMovItem, rsBS!cCtaContCod, Abs(Round(rsBS!P_AFBND, 2)) * -1
                    Else
                        oMov.InsertaMovCta rs!nMovNro, rsBS!nMovItem, rsBS!cCtaContCod, Abs(Round(rsBS!monto, 2)) * -1
                    End If
                Else
                    oMov.InsertaMovCta rs!nMovNro, rsBS!nMovItem, rsBS!cCtaContCod, Abs(Round(rsBS!monto, 2)) * -1
                End If
                
                lnItem = rsBS!nMovItem
                rsBS.MoveNext
            Wend
            
            rsBS.MoveFirst
            
            If lnItem = rsBS.RecordCount Then
                lnItem = 0
            End If
            
            While Not rsBS.EOF
                lsCtaTemp = ""
                lsCtaCnt = GetOpeCtaCta(rs!COPECOD, "", rsBS!cCtaContCod, Format(rsBS!nAlmacenCod, "00"))
                lsCtaTemp = GetCtaOpeBS(rs!COPECOD, rsBS!cBSCod)
                If lsCtaTemp <> "" Then
                    lsCtaCnt = lsCtaTemp
                End If
                
                If lsCtaCnt = "" Then
                    MsgBox "hola"
                End If
                
                lsCtaCnt = Replace(lsCtaCnt, "AG", GetCtaAreaAge(rs!cAreaCod, rs!cAgeCod))
                
                If Left(rsBS!cBSCod, 3) = "112" Or Left(rsBS!cBSCod, 3) = "113" Then
                    If rsBS!P_AFBND <> 0 Then
                        oMov.InsertaMovCta rs!nMovNro, rsBS.RecordCount + lnItem + rsBS!nMovItem, lsCtaCnt, Abs(Round(rsBS!P_AFBND, 2))
                    Else
                        oMov.InsertaMovCta rs!nMovNro, rsBS.RecordCount + lnItem + rsBS!nMovItem, lsCtaCnt, Abs(Round(rsBS!monto, 2))
                    End If
                Else
                    oMov.InsertaMovCta rs!nMovNro, rsBS.RecordCount + lnItem + rsBS!nMovItem, lsCtaCnt, Abs(Round(rsBS!monto, 2))
                End If
                
                rsBS.MoveNext
            Wend
            
            
            Set oAsiento = New NContImprimir
            
            If lsAsientos = "" Then
                lsAsientos = oAsiento.ImprimeAsientoContable(oMov.GetcMovNro(rs!nMovNro), 66, 120)
            Else
                lsAsientos = lsAsientos & oAsiento.ImprimeAsientoContable(oMov.GetcMovNro(rs!nMovNro), 66, 120) & oImpresora.gPrnSaltoPagina
            End If
        End If
        
        rs.MoveNext
        Me.PrgBar.value = Me.PrgBar.value + 1
    Wend
    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    
    lsAsientos = CON & lsAsientos
    MsgBox "Proceso Finalizado!" + Space(10), vbInformation, "Aviso"
    fraBarra.Visible = False
    CmdSalir.Visible = True
    oPrevio.Show lsAsientos, Caption, True
    Caption = lsCadena
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'------------------------------------------
'Para verificar la existencia de saldos
'------------------------------------------
Function YaHaySaldos(pnMes As Integer, pnAnio As Integer, pnAlmacen As Integer) As Boolean
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

YaHaySaldos = False
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet("select top 1 cBSCod from BSSaldos where month(dSaldo)=" & pnMes & " and year(dSaldo)=" & pnAnio & " and nAlmCod= " & pnAlmacen & " ")
   If Not rs.EOF Then
      YaHaySaldos = True
   End If
End If
End Function





