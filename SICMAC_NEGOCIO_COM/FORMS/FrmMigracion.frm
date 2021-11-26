VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMigracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Migracion SIMACI-SIAF"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   6870
      TabIndex        =   16
      Top             =   5850
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   6315
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin VB.CommandButton CmdMigracion 
         Caption         =   "Migracion"
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   5850
         Width           =   1245
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   405
         Left            =   6870
         TabIndex        =   10
         Top             =   1110
         Width           =   1335
      End
      Begin VB.ListBox LstCredito 
         Height          =   1035
         Left            =   5040
         TabIndex        =   9
         Top             =   870
         Width           =   1725
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos"
         Height          =   1155
         Left            =   90
         TabIndex        =   5
         Top             =   750
         Width           =   4935
         Begin SICMACT.TxtBuscar TxtBuscar1 
            Height          =   345
            Left            =   930
            TabIndex        =   17
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   609
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   690
            Width           =   3825
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   90
            TabIndex        =   7
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   255
            Left            =   90
            TabIndex        =   6
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Migracion"
         Height          =   585
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   4905
         Begin VB.CheckBox ChkMovimientos 
            Caption         =   "Movimientos"
            Height          =   195
            Left            =   3360
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox ChkCalendario 
            Caption         =   "Calendario"
            Height          =   195
            Left            =   1890
            TabIndex        =   3
            Top             =   240
            Width           =   1515
         End
         Begin VB.CheckBox ChkLineaCredito 
            Caption         =   "Linea de Credito"
            Height          =   195
            Left            =   210
            TabIndex        =   2
            Top             =   240
            Width           =   1545
         End
      End
      Begin MSComctlLib.ListView LstReportes 
         Height          =   2565
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Seleccione un Reporte para su impresion"
         Top             =   3180
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483647
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Vencimiento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha de Pago"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Aplicado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Capital"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Interes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Mora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Capital Pagado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Interes Pagado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Mora Pagada"
            Object.Width           =   2540
         EndProperty
      End
      Begin SICMACT.TxtBuscar txtBuscarLinea 
         Height          =   345
         Left            =   1470
         TabIndex        =   19
         Top             =   2490
         Width           =   1545
         _ExtentX        =   3149
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         EnabledText     =   0   'False
      End
      Begin VB.Label lblLineaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3060
         TabIndex        =   20
         Top             =   2490
         Width           =   4515
      End
      Begin VB.Label Lss 
         AutoSize        =   -1  'True
         Caption         =   "Linea de Credito"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Calendario Vigente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   2850
         Width           =   2715
      End
      Begin VB.Label LblNroCredito 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1470
         TabIndex        =   12
         Top             =   1980
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro del Credito:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2070
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsBaseCred As String
Dim gLineaCodigo As String
Dim RLinea As ADODB.Recordset
Dim gcCtacod As String
Dim dbBase As ADODB.Connection


Private Sub ChkLineaCredito_Click()
    If ChkLineaCredito.value = vbChecked Then
        txtBuscarLinea.Enabled = True
        lblLineaDesc.Caption = ""
    Else
        txtBuscarLinea.Text = ""
        lblLineaDesc.Caption = ""
        lblLineaDesc.ToolTipText = ""
    End If
End Sub

Private Sub CmdMigracion_Click()
    If MsgBox("Desea migrar los calendarios", vbInformation + vbYesNo, "AVISO") = vbYes Then
        If ChkLineaCredito.value = vbUnchecked And ChkCalendario.value = vbUnchecked And _
           chkMovimientos.value = vbUnchecked Then
             MsgBox "Debe seleccionar una de la opciones para la migracion", vbInformation, "AVISO"
             Exit Sub
            
        Else
            sCodCta = InputBox("Ingrese Cuenta del Siaf", "SIAF")
            If ChkLineaCredito.value = Checked Then
                MigracionLineas sCodCta, gcCtacod
                MigracionCalendario sCodCta, gcCtacod
            End If
            If ChkCalendario.value = Checked And ChkLineaCredito.value = vbUnchecked Then
                MigracionCalendario sCodCta, gcCtacod
            End If
            If chkMovimientos.value = Checked Then
                
            End If
        End If
    End If
End Sub
Sub MigracionLineas(ByVal pscCodCta As String, ByVal pscCtaCod As String)
    Dim lcLineaCredito As String
    Dim nMontoCol As Double
    Dim nTasaInteresCompesatoria As Integer
    Dim nTasaMoratoria As Integer
    Dim nTasaGracia As Integer
    
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    ' se actualiza la linea antigua y la nueva linea de ColocLineaCreditoSaldo
    Call DeterminaRutas(Mid(pscCtaCod, 4, 2))
        
    sSql = "Select cLineaCred,nMontoCol From Colocaciones Where cCtaCod='" & pscCtaCod & "'"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        lcLineaCredito = rs!cLineaCred
        nMontoCol = rs!nMontoCol
    End If
    Set rs = Nothing
    
    If lcLineaCredito <> "" Then
        ' actualizamos el saldo de la linea antigua
        Set oConec = New DConecta
        oConec.AbreConexion
        
        oConec.ConexionActiva.BeginTrans
        
        sSql = "Update ColocLineaCreditoSaldo"
        sSql = sSql & " set nMontoColocado=nMontoColocado - " & nMontoCol & " ,"
        sSql = sSql & " nSaldoCap = nSaldoCap + " & nMontoCol
        sSql = sSql & " Where cLineaCred='" & lcLineaCredito & " '"
        
        oConec.ConexionActiva.Execute sSql
        
        ' actualizamos el saldo de la linea nueva
        
        sSql = "Update ColocLineaCreditoSaldo"
        sSql = sSql & " set nMontoColocado=nMontoColocado - " & nMontoCol & " ,"
        sSql = sSql & " nSaldoCap = nSaldoCap + " & nMontoCol
        sSql = sSql & " Where cLineaCred='" & gLineaCodigo & " '"
        
        oConec.ConexionActiva.Execute sSql
        
        
        ' actualizamos las colocaciones
        
        sSql = "Update Colocaciones "
        sSql = sSql & " set cLineaCred='" & gLineaCodigo & "'"
        sSql = sSql & " Where cCtaCod='" & pscCtaCod & "'"
        
        oConec.ConexionActiva.Execute sSql
        
        'actualizamos la tabla productotasainteres
        sSql = "Select  nColocLinCredTasaTpo,nTasaIni"
        sSql = sSql & " From ColocLineaCreditoTasa"
        sSql = sSql & " Where cLineaCred='" & gLineaCodigo & "'"
        
        Set rs = oConec.CargaRecordSet(sSql)
        
        Do Until rs.EOF
            If rs!nColocLinCredTasaTpo = "1" Then
                sSql = "Update ProductoTasaInteres "
                sSql = sSql & " Set nTasaInteres=" & rs!nTasaIni
                sSql = sSql & "Where cCtaCod='" & pscCtaCod & "' and nPrdTasaInteres=1"
                oConec.ConexionActiva.Execute sSql
                
                sSql = "Update Producto"
                sSql = sSql & " Set nTasaInteres=" & rs!nTasaIni
                sSql = sSql & "Where cCtaCod='" & pscCtaCod & "'"
                oConec.ConexionActiva.Execute sSql
                
            ElseIf rs!nColocLinCredTasaTpo = "3" Then
                sSql = "Update ProductoTasaInteres "
                sSql = sSql & " Set nTasaInteres=" & rs!nTasaIni
                sSql = sSql & "Where cCtaCod='" & pscCtaCod & "' and nPrdTasaInteres=3"
                oConec.ConexionActiva.Execute sSql
            ElseIf rs!nColocLinCredTasaTpo = "5" Then
                sSql = "Update ProductoTasaInteres "
                sSql = sSql & " Set nTasaInteres=" & rs!nTasaIni
                sSql = sSql & "Where cCtaCod='" & pscCtaCod & "' and nPrdTasaInteres=5"
                oConec.ConexionActiva.Execute sSql
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
        
        oConec.ConexionActiva.CommitTrans
        
        MsgBox "Se actualizo correctamente las lineas", vbInformation, "AVISO"
    Else
        MsgBox "No se encontro linea de credito original", vbInformation, "AVISO"
    End If
    Exit Sub
ErrHandler:
    oConec.ConexionActiva.RollbackTrans
    If oConec.ConexionActiva.State = 1 Then
        oConec.ConexionActiva.Close
    End If
    Set oConec = Nothing
    MsgBox "Se ha producido un error " & Err.Description, vbInformation, "AVISO"
End Sub

Sub MigracionCalendario(ByVal pscCodCta As String, ByVal pscCtaCod As String)
    Dim oConec As DConecta
    Dim sSql As String
    Dim nCuota As Integer
    Dim nMonto As Double
    Dim nMontoPagado As Double
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Call DeterminaRutas(Mid(pscCtaCod, 4, 2))
    
    ' vaciando la datos al SQL
     VaciandoDataFoxSQL pscCodCta
    
    ' Se elimina el Calendario del SIMACT
    Set oConec = New DConecta
    oConec.AbreConexion
    
    oConec.BeginTrans
    sSql = "Delete From ColocCalendDet where cCtaCod='" & pscCtaCod & "'"
    oConec.ConexionActiva.Execute sSql
    
    sSql = "Delete From ColocCalendario where cCtaCod='" & pscCtaCod & "'"
    oConec.ConexionActiva.Execute sSql
    
    ' insertando en ColocCalendario
    sSql = "Select '" & pscCtaCod & "' as cCtaCod,"
    sSql = sSql & " 3 as nNroCalen,"
    sSql = sSql & " dbcmacicamig.dbo.fn_nColocCalenApl(c.ctipope) as nColocCalendApl,"
    sSql = sSql & " Cast(c.cnrocuo as int) as nCuota,"
    sSql = sSql & " c.dfecven as dVenc,"
    sSql = sSql & " c.dfecpag as dPago,"
    sSql = sSql & " dbcmacicamig.dbo.fn_nColocCalendEstadoxx(c.cestado,c.ncapita,c.ncappag) as nColocCalendEstado,"
    sSql = sSql & " 'Migracion' as cDescripcion,"
    sSql = sSql & " NULL as cColocCalendFlag,"
    sSql = sSql & " isnull(dbcmacicamig.dbo.fn_nCalendProc(c.cestado),2) as nCalendProc,"
    sSql = sSql & " null As cColocMiVivEval"
    sSql = sSql & " From  dbcmacicafox.dbo.kpydppg c"
    sSql = sSql & " Where C.ccodcta='" & pscCodCta & "'"

    Set rs = oConec.CargaRecordSet(sSql)
    
    Do Until rs.EOF
        sSql = "Insert into ColocCalendario Values('" & rs!cCtaCod & "'," & rs!nNroCalen & ","
        sSql = sSql & rs!nColocCalendApl & "," & rs!nCuota & ",'" & Format(rs!dVenc, "MM/dd/yyyy") & "','" & Format(rs!dPago, "mm/dd/yyyy") & "',"
        sSql = sSql & rs!nColocCalendEstado & ",'" & rs!cDescripcion & "','" & rs!cColocCalendFlag & "',"
        sSql = sSql & IIf(IsNull(rs!nCalendProc), Null, rs!nCalendProc) & ",'" & rs!cColocMiVivEval & "')"
        
        oConec.ConexionActiva.Execute sSql
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    ' insertando Detalle de capital de Calendario
    
    sSql = "select '" & pscCtaCod & "' as cCtaCod,"
    sSql = sSql & " 3 as nNroCalen,"
    sSql = sSql & " dbcmacicamig.dbo.fn_nColocCalenApl(c.cTipOpe) as nColocCalendApl,"
    sSql = sSql & " cast(c.cnrocuo as int) as nCuota,"
    sSql = sSql & " 1000 as nPrdConceptoCod,"
    sSql = sSql & " c.ncapita as nMonto,"
    sSql = sSql & " c.ncappag as nMontoPagado,"
    sSql = sSql & " '' as cFlag"
    sSql = sSql & " From  dbcmacicafox..kpydppg C"
    sSql = sSql & " Where C.ccodcta='" & pscCodCta & "'"

    Set rs = oConec.CargaRecordSet(sSql)
    
    Do Until rs.EOF
    sSql = "Insert into ColocCalendDet Values('" & rs!cCtaCod & "'," & rs!nNroCalen & ","
    sSql = sSql & rs!nColocCalendApl & "," & rs!nCuota & "," & rs!nPrdConceptoCod & ","
    sSql = sSql & rs!nMonto & "," & rs!nMontoPagado & ",'" & rs!cFlag & "')"
    
    oConec.ConexionActiva.Execute sSql
    rs.MoveNext
    Loop
    Set rs = Nothing
    ' insertando el detalle de interes
    
    sSql = "select '" & pscCtaCod & "' as cCtaCod,"
    sSql = sSql & " 3 as nNroCalen,"
    sSql = sSql & " dbcmacicamig.dbo.fn_nColocCalenApl(c.cTipOpe) as nColocCalendApl,"
    sSql = sSql & " cast(c.cnrocuo as int) as nCuota,"
    sSql = sSql & " 1100 as nPrdConceptoCod,"
    sSql = sSql & " c.nintere as nMonto,"
    sSql = sSql & " c.nintpag as nMontoPagado,"
    sSql = sSql & " '' as cFlag"
    sSql = sSql & " From  dbcmacicafox..kpydppg c"
    sSql = sSql & " Where c.cCodCta='" & pscCodCta & "'"

    Set rs = oConec.CargaRecordSet(sSql)

    Do Until rs.EOF
        sSql = "Insert into ColocCalendDet Values('" & rs!cCtaCod & "'," & rs!nNroCalen & ","
        sSql = sSql & rs!nColocCalendApl & "," & rs!nCuota & "," & rs!nPrdConceptoCod & ","
        sSql = sSql & rs!nMonto & "," & rs!nMontoPagado & ",'" & rs!cFlag & "')"
        
        oConec.ConexionActiva.Execute sSql
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    ' insertando el detalle de mora
        
    sSql = " select '" & pscCtaCod & "' as cCtaCod,"
    sSql = sSql & " 3 as nNroCalen,"
    sSql = sSql & " dbcmacicamig.dbo.fn_nColocCalenApl(c.cTipOpe) as nColocCalendApl,"
    sSql = sSql & " cast(c.cnrocuo as int) as nCuota,"
    sSql = sSql & " 1101 as nPrdConceptoCod,"
    sSql = sSql & " c.nintmor + c.nmorpag  as nMonto,"
    sSql = sSql & " c.nmorpag as nMontoPagado,"
    sSql = sSql & " '' as cFlag"
    sSql = sSql & " from  dbcmacicafox..kpydppg c"
    sSql = sSql & " Where c.cCodCta='" & pscCodCta & "'"
    
    Set rs = oConec.CargaRecordSet(sSql)
    
    Do Until rs.EOF
        sSql = "Insert into ColocCalendDet Values('" & rs!cCtaCod & "'," & rs!nNroCalen & ","
        sSql = sSql & rs!nColocCalendApl & "," & rs!nCuota & "," & rs!nPrdConceptoCod & ","
        sSql = sSql & rs!nMonto & "," & rs!nMontoPagado & ",'" & rs!cFlag & "')"
        
        oConec.ConexionActiva.Execute sSql
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    
    
    oConec.CommitTrans
    Set oConec = Nothing
    
    MsgBox "Se genero los calendarios correctamente", vbInformation, "AVISO"
    Exit Sub
ErrHandler:
        oConec.ConexionActiva.RollbackTrans
        Set oConec = Nothing
        MsgBox "Error en los calendarios", vbInformation, "AVISO"
End Sub

Sub VaciandoDataFoxSQL(ByVal pscCodCta)
    Dim sSql As String
    Dim sSql1 As String
    Dim rs As ADODB.Recordset
    Dim oConec As DConecta
    
    On Error GoTo ErrHandler
    AbreConexionFox
    sSql = "Select * From " & gsBaseCred & "kpydppg "
    sSql = sSql & " Where cCodCta='" & pscCodCta & "'"
    
    Set rs = CargaRecordFox(sSql)
    sSql1 = ""
    Do Until rs.EOF
        sSql1 = sSql1 & " Insert into dbcmacicafox..kpydppg Values('" & rs!cCodCta & "','" & Format(rs!dfecven, "MM/dd/yyyy") & "','" & Format(rs!dfecpag, "MM/dd/yyyy") & "','"
        sSql1 = sSql1 & rs!cEstado & "','" & rs!cTipOpe & "','" & rs!cnrocuo & "'," & rs!ncapita & ","
        sSql1 = sSql1 & rs!nintere & "," & rs!nCapPag & "," & rs!nIntPag & "," & rs!nIntMor & "," & rs!nmorpag & ","
        sSql1 = sSql1 & rs!NIntDif & "," & rs!nIntdPag & "," & rs!NotrPag & ",'" & rs!cCodUsu & "','" & Format(rs!dFecMod, "MM/dd/yyyy") & "','"
        sSql1 = sSql1 & rs!cFlag & "'," & rs!nCofide & "," & rs!NCofpag & "," & rs!NMorCof & "," & rs!NmCofpg & ","
        sSql1 = sSql1 & rs!NDiaAtr & ")"
        rs.MoveNext
    Loop
    
    CierraConexionFox
    Set rs = Nothing
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.BeginTrans
    sSql = "Delete From dbcmacicafox..kpydppg where ccodcta='" & pscCodCta & "'"
    oConec.ConexionActiva.Execute sSql
    oConec.ConexionActiva.Execute sSql1
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Sub
ErrHandler:
    oConec.ConexionActiva.RollbackTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Sub MigracionMovimientos(ByVal pscCodCta As String, ByVal pscCtaCod As String)

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub LstCredito_DblClick()
    Dim cCtaCod As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    
    LstReportes.ListItems.Clear
    cCtaCod = LstCredito.Text
    gcCtacod = cCtaCod
    If cCtaCod <> "" Then
        
     sSql = "Select CC.nCuota,CC.dVenc,CC.dPago,"
     sSql = sSql & " Tipo=Case CC.nColocCalendApl"
     sSql = sSql & " When 1 Then 'Pago'"
     sSql = sSql & " When 0 Then 'Desembolso'"
     sSql = sSql & " End,"
     sSql = sSql & " Capital=isnull((Select nMonto From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1000),0),"
     sSql = sSql & " Interes=isnull((Select nMonto From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1100),0),"
     sSql = sSql & " Mora=isnull((Select nMonto From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1001),0),"
     sSql = sSql & " CapitalPagado=isnull((Select nMontoPagado From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1000),0),"
     sSql = sSql & " InteresPagado=isnull((Select nMontoPagado From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1100),0),"
     sSql = sSql & " MoraPagado=isnull((Select nMontoPagado From ColocCalendDet CD Where CD.cCtaCod=CC.cCtaCod and"
     sSql = sSql & " CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
     sSql = sSql & " CC.nCuota=CD.nCuota and CD.nPrdConceptoCod=1001),0)"
     sSql = sSql & " From ColocCalendario CC"
     sSql = sSql & " Where CC.cCtaCod='" & cCtaCod & "' and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendario"
     sSql = sSql & " Where cCtaCod='" & cCtaCod & "')"
     sSql = sSql & " Order By CC.nCuota,CC.nColocCalendApl"
     
     Set oConec = New DConecta
     oConec.AbreConexion
     Set rs = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
     
     Do Until rs.EOF
        Set Item = LstReportes.ListItems.Add(, , rs!nCuota)
        Item.SubItems(1) = Format(rs!dVenc, "dd/mm/yyyy")
        Item.SubItems(2) = Format(IIf(IsNull(rs!dPago), "", rs!dPago), "dd/mm/yyyy")
        Item.SubItems(3) = rs!Tipo
        Item.SubItems(4) = Format(rs!Capital, "#0.00")
        Item.SubItems(5) = Format(rs!Interes, "#0.00")
        Item.SubItems(6) = Format(rs!Mora, "#0.00")
        Item.SubItems(7) = Format(rs!CapitalPagado, "#0.00")
        Item.SubItems(8) = Format(rs!InteresPagado, "#0.00")
        Item.SubItems(9) = Format(rs!MoraPagado, "#0.00")
        rs.MoveNext
     Loop
    End If
End Sub

'Dim Mat_Creditos() As Double
Private Sub TxtBuscar1_EmiteDatos()
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
  '  Dim nMaxIndice As Integer
    
    'Cargando a la persona
    sSql = "Select cPersNombre From Persona Where cPersCod='" & TxtBuscar1.Text & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        txtNombre = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
    End If
    Set rs = Nothing
    
    'cargando lista de creditos vigentes en listbox
    If txtNombre <> "" Then
    
        sSql = "Select P.cCtaCod as cCtaCod"
        sSql = sSql & " From ProductoPersona PP"
        sSql = sSql & " Inner Join Producto P on PP.cCtaCod=P.cCtaCod"
        sSql = sSql & " Where PP.cPersCod='" & TxtBuscar1.Text & "' and PP.nPrdPersRelac=20 and P.nPrdEstado=2020"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        
        Do Until rs.EOF
            'nMaxIndice = 0
          '  nMaxIndice = LBound(Mat_Creditos, 1)
           ' ReDim Preserve Mat_Creditos(nMaxIndice + 1)
            LstCredito.AddItem rs!cCtaCod
            'LstCredito.ItemData(LstCredito.NewIndex) = CInt(rs!cCtaCod)
            rs.MoveNext
        Loop
        Set rs = Nothing
    End If
    
End Sub

Sub AbreConexionFox()
On Error GoTo ErrorConex

    psConexion = "Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=F:\APL\TCOS\;SourceType=DBF;Exclusive=No;Collate=GENERAL"
    
    Set dbBase = New ADODB.Connection
    dbBase.Open psConexion
    dbBase.CommandTimeout = 7200
    dbBase.Execute "SET DELETE ON"
    
    Exit Sub
ErrorConex:
    MsgBox "Error:" & Err.Description & " [" & Err.Number & "] ", vbInformation, "Aviso"
    End
End Sub
Function CargaRecordFox(ByVal Sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Sql, dbBase, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set CargaRecordFox = rs
    rs.ActiveConnection = Nothing
End Function
Public Sub CierraConexionFox()
  'If nSalida = 1 Then
     If dbBase Is Nothing Then Exit Sub
     If dbBase.State = adStateOpen Then
        dbBase.Close
        Set dbBase = Nothing
     End If
  'End If
End Sub
Sub DeterminaRutas(ByVal cCodAgencias As String)
    Select Case Trim(cCodAgencias)
        Case "01" 'ica
            gsBaseCred = "F:\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.ICA\"
'            gsBaseAho = "F:\APL.FIN\AHO.ICA\"
'            gsBaseKPR = "F:\APL.FIN\KPR.ICA\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.ICA\"
        Case "02" 'CAÑETE
            gsBaseCred = "F:\AGENCIAS\CANETE\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.CAN\"
'            gsBaseAho = "F:\APL.FIN\AHO.CAN\"
'            gsBaseKPR = "F:\APL.FIN\KPR.CAN\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.CAN\"
        Case "03" 'CHINCHA
            gsBaseCred = "F:\AGENCIAS\CHINCHA\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.CHI\"
'            gsBaseAho = "F:\APL.FIN\AHO.ICA\"
'            gsBaseKPR = "F:\APL.FIN\KPR.ICA\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.CHI\"
        Case "04" 'NASCA
            gsBaseCred = "F:\AGENCIAS\NASCA\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.NAS\"
'            gsBaseAho = "F:\APL.FIN\AHO.NAS\"
'            gsBaseKPR = "F:\APL.FIN\KPR.NAS\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.NAS\"
        Case "05" 'PUQUIO
            gsBaseCred = "F:\AGENCIAS\PUQUIO\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.PUQ\"
'            gsBaseAho = "F:\APL.FIN\AHO.PUQ\"
'            gsBaseKPR = "F:\APL.FIN\KPR.PUQ\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.PUQ\"
        Case "06" 'HUAMANGA
            gsBaseCred = "F:\AGENCIAS\HUAMANGA\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.HUA\"
'            gsBaseAho = "F:\APL.FIN\AHO.HUA\"
'            gsBaseKPR = "F:\APL.FIN\KPR.HUA\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.HUA\"
        Case "07" 'MALA
            gsBaseCred = "F:\AGENCIAS\MALA\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.MAL\"
'            gsBaseAho = "F:\APL.FIN\AHO.MAL\"
'            gsBaseKPR = "F:\APL.FIN\KPR.MAL\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.MAL\"
        Case "08" 'PALPA
            gsBaseCred = "F:\AGENCIAS\PALPA\APL\KPY\"
'            gsBaseCli = "F:\APL.FIN\CLI.PAL\"
'            gsBaseAho = "F:\APL.FIN\AHO.PAL\"
'            gsBaseKPR = "F:\APL.FIN\KPR.PAL\"
'            gsBaseTCOS = "F:\APL.FIN\TCOS.PAL\"
        Case "09" 'IMPERIAL
            gsBaseCred = "F:\AGENCIAS\IMPERIAL\APL\KPY\"
    End Select
End Sub


Private Sub txtBuscarLinea_EmiteDatos()
Dim sCodigo As String
Dim oLineas As dLineaCredito
Dim RLineaProducto As ADODB.Recordset

If ChkLineaCredito.value = vbChecked Then
    Set oLineas = New dLineaCredito
    Set RLineaProducto = oLineas.RecuperaLineasProductoArbol(Mid(gcCtacod, 6, 3), Mid(gcCtacod, 9, 1))
    Set oLineas = Nothing
            
    txtBuscarLinea.rs = RLineaProducto
    Set RLineaProducto = Nothing
            
    'txtBuscarLinea.Text = ""
    
    sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
    
    If sCodigo <> "" Then
    '    txtBuscarLinea.Text = sCodigo
        If txtBuscarLinea.psDescripcion <> "" Then lblLineaDesc = txtBuscarLinea.psDescripcion Else lblLineaDesc = ""
            'VERIFICAR
           'Carga Datos de la Linea de Credito seleccionada
           Set oLineas = New dLineaCredito
           Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
           Set oLineas = Nothing
           If RLinea.RecordCount > 0 Then
              Call CargaDatosLinea
           Else
              MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
              txtBuscarLinea.Text = ""
              lblLineaDesc = ""
           End If
        
        'txtBuscarLinea.Text = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
        
    Else
        lblLineaDesc = ""
    End If
    
    gLineaCodigo = sCodigo
    txtBuscarLinea.Enabled = True
End If
End Sub

Private Sub CargaDatosLinea()
    Dim nTasaCompesatoria As Double
    Dim nTasaGracia As Double
    Dim nTasaMoratoria As Double
    
    If Trim(txtBuscarLinea.Text) = "" Then
        Exit Sub
    End If

    If RLinea!nTasaIni <> RLinea!nTasafin Then
        nTasaCompesatoria = Format(RLinea!nTasaIni, "#0.0000")
    End If
    If RLinea!nTasaGraciaIni <> RLinea!nTasaGraciaFin Then
        nTasaGracia = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
    End If
    
    If RLinea!nTasaMoraIni <> RLinea!nTasaMoraFin Then
        nTasaMoratoria = Format(IIf(IsNull(RLinea!nTasaMoraIni), 0, RLinea!nTasaMoraIni), "#0.0000")
    End If
    lblLineaDesc.ToolTipText = "Tasa Compesatoria: " & CStr(nTasaCompesatoria) & vbCrLf & _
                               "Tasa Gracia: " & CStr(nTasaGracia) & vbCrLf & _
                               "Tasa Moratoria:" & CStr(nTasaMoratoria)
End Sub

