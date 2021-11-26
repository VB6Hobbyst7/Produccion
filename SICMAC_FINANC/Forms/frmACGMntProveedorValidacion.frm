VERSION 5.00
Begin VB.Form frmACGMntProveedorValidacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Resultados de Validacion"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmACGMntProveedorValidacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   5190
      Width           =   885
   End
   Begin VB.CommandButton cmdResumen 
      Caption         =   "&Resumen"
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
      Height          =   375
      Left            =   5250
      TabIndex        =   12
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalculadora 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Calculadora"
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
      Height          =   375
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5190
      Width           =   1395
   End
   Begin VB.CommandButton cmdExtornoArchivo 
      Caption         =   "<< &Extorno Archivo >>"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   5190
      Width           =   2070
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Height          =   375
      Left            =   6870
      TabIndex        =   5
      Top             =   5190
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   10035
      Begin VB.Frame Frame2 
         Height          =   885
         Left            =   210
         TabIndex        =   14
         Top             =   1035
         Visible         =   0   'False
         Width           =   5265
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Logisitca"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   17
            Top             =   120
            Width           =   1125
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "x Agencia"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   16
            Top             =   150
            Width           =   1065
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin Sicmact.TxtBuscar txtAge 
            Height          =   345
            Left            =   900
            TabIndex        =   18
            Top             =   450
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Appearance      =   0
            BackColor       =   14811132
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
            EnabledText     =   0   'False
         End
         Begin VB.Label lblAgencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1920
            TabIndex        =   20
            Top             =   450
            Width           =   3255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   510
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5295
         TabIndex        =   6
         Top             =   315
         Width           =   1065
      End
      Begin Sicmact.TxtBuscar txtProveedor 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   285
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   609
         Appearance      =   0
         BackColor       =   14811132
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
         EnabledText     =   0   'False
      End
      Begin VB.Frame FraFechaMov 
         Height          =   525
         Left            =   6600
         TabIndex        =   8
         Top             =   120
         Width           =   2550
         Begin VB.TextBox txtFFiltro 
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
            Height          =   300
            Left            =   1275
            MaxLength       =   8
            TabIndex        =   10
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Filtro:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   90
            TabIndex        =   9
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Doc Enviado"
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
         Height          =   180
         Left            =   135
         TabIndex        =   21
         Top             =   360
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "&Registrar"
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
      Left            =   7980
      TabIndex        =   1
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
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
      Left            =   9105
      TabIndex        =   0
      Top             =   5190
      Width           =   1095
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   4245
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7488
      Cols0           =   27
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmACGMntProveedorValidacion.frx":08CA
      EncabezadosAnchos=   "400-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-0-1500-0-1400-900-2500-2500-2300-2300-1200-1200-2000-1800"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-17-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-C-C-C-R-C-R-L-L-L-R-C-C-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-2-0-2-0-0-0-2-0-0-4-0-0"
      TextArray0      =   "Nro."
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   360
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmACGMntProveedorValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lsTipoB As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub chkTodos_Click()
 Dim i As Integer
    If Me.chkTodos.value = 1 Then
        For i = 1 To Me.fg.Rows - 1
            Me.fg.TextMatrix(i, 2) = 1
        Next i
    Else
        For i = 1 To Me.fg.Rows - 1
            Me.fg.TextMatrix(i, 2) = 0
        Next i
    End If
End Sub

Private Sub cmdCalculadora_Click()
     Shell "calc.exe", vbMaximizedFocus
End Sub

Private Sub cmdExtornoArchivo_Click()
    Dim i       As Integer
    Dim lbBan   As Boolean
    Dim oCon    As New DConecta
    Dim sSql    As String
    
    lbBan = False
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 19) <> "" Then
            lbBan = True
            Exit For
        End If
    Next i
    
    
    lbBan = False
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            If Me.fg.TextMatrix(i, 20) <> "" Then
                lbBan = True
            End If
        End If
    Next i
    
    If lbBan Then
        MsgBox "Alguno de los registros marcados ya fue pagado.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If Not lbBan Then
        If MsgBox("Desea Extornar el Archivo Generado ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        oCon.AbreConexion
        For i = 1 To Me.fg.Rows - 1
            If Me.fg.TextMatrix(i, 2) = "." Then
                sSql = " Update movControlPagoSunat set bVigente=0 ,bValido = 0 where nMovNro=" & Me.fg.TextMatrix(i, 10) & " and bvigente=1"
                oCon.Ejecutar sSql
            End If
        Next i
        oCon.CierraConexion
        MsgBox "Extorno realizado con exito.", vbInformation, "Aviso"
        Me.txtProveedor.Text = ""
        cmdProcesar_Click
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Extorno Operación "
                Set objPista = Nothing
                '****
    Else
        MsgBox "No puede extornar un archivo donde ya se han registado resultados.", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdImprimir_Click()
    Call GenerarExcel
End Sub

Private Sub cmdProcesar_Click()
    Me.fg.Clear
    Me.fg.Rows = 2
    Me.fg.FormaCabecera
    Me.txtProveedor.Text = ""
    Me.txtProveedor.rs = CargaObjeto
End Sub

Private Sub cmdRegistrar_Click()
    Dim sSql    As String
    Dim rs      As New ADODB.Recordset
    Dim oCon    As New DConecta
    Dim i       As Integer
    Dim oMov    As New DMov
    Dim lbError As Boolean
    
    lbError = False
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            lbError = True
        End If
    Next i
    
    If lbError = False Then
        MsgBox "No ha marcado ningun registro.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    lbError = False
    fg.col = 17
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            If CDbl(Me.fg.TextMatrix(i, 15)) > 0 Then
                If CDbl(Me.fg.TextMatrix(i, 17)) >= CDbl(Me.fg.TextMatrix(i, 15)) Then
                    lbError = True
                    fg.row = i
                End If
            End If
        End If
    Next i
    
    If lbError Then
        MsgBox "En alguno de los registros marcados consigna un monto de cobranza mayor o igual al monto a pagar.", vbCritical, "Aviso"
        fg.SetFocus
        Exit Sub
    End If
    
    
    lbError = False
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            If Me.fg.TextMatrix(i, 20) <> "" Then
                lbError = True
            End If
        End If
    Next i
    
    If lbError Then
        MsgBox "Alguno de los registros marcados ya fue pagado.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If MsgBox(" ¿ Esta seguro de registrar el resultado ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    oCon.AbreConexion
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            sSql = " Update movControlPagoSunat set cMovNroRes='" & Trim(oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)) & "' ,nImporteCoactivo=" & CDbl(IIf(Me.fg.TextMatrix(i, 17) = "", 0, Me.fg.TextMatrix(i, 17))) & " where nMovNro=" & Me.fg.TextMatrix(i, 10) & " and bvigente=1 and bvalido=1"
            oCon.Ejecutar sSql
        End If
    Next i
    oCon.CierraConexion
    
    MsgBox "Datos Registrados", vbInformation, "Aviso"
    txtProveedor_EmiteDatos
                
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Registro Operación "
                Set objPista = Nothing
                '****
End Sub

Private Sub cmdResumen_Click()
    GenerarExcel ("R")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fg_OnCellChange(pnRow As Long, pnCol As Long)
    If Me.fg.col = 17 Then
        If Me.fg.TextMatrix(Me.fg.row, 15) < 0 And Me.fg.TextMatrix(Me.fg.row, 17) > 0 Then
            MsgBox " El monto retenido no valido ", vbInformation, "Aviso"
            Me.fg.TextMatrix(Me.fg.row, 17) = "0.00"
            Exit Sub
        End If
    End If
End Sub

Private Sub fg_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    fg.row = pnRow
    fg.col = 17
End Sub

Private Sub Form_Load()
    Dim oAge As New DActualizaDatosArea
   
    Me.txtFFiltro.Text = Format(gdFecSis, gsFormatoMovFecha)
    Me.cmdCalculadora.Enabled = True
    Me.cmdResumen.Enabled = True
    Me.txtAge.rs = oAge.GetAgencias
    
    Me.txtAge.Text = ""
    Me.lblAgencia.Caption = ""
    Me.txtAge.Enabled = False
    lsTipoB = "T"
End Sub

Public Function CargaObjeto(Optional psPagados As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaObjetoErr
   Dim oCon As New DConecta
   Dim psSql As String
   
   If oCon.AbreConexion Then
      If lsTipoB = "T" Then
        psSql = "select distinct cmovnro,'PROVEEDOR'+ CMOVNRO,len(cmovnro) from movcontrolpagosunat Where cmovnro Like '" & Me.txtFFiltro.Text & "%' and bvigente=1"
        
      ElseIf lsTipoB = "A" Then
        psSql = "select distinct cmovnro,'PROVEEDOR'+ CMOVNRO,len(cmovnro) from movcontrolpagosunat MS "
        psSql = psSql & " INNER Join MovProvisionAgencia MPA on MPA.nmovnro=ms.nmovnro and cAgeCod='" & Me.txtAge.Text & "'"
        psSql = psSql & " Where cmovnro Like '" & Me.txtFFiltro.Text & "%' and bvigente=1"
        
      ElseIf lsTipoB = "L" Then
        psSql = "select distinct cmovnro,'PROVEEDOR'+ CMOVNRO,len(cmovnro) from movcontrolpagosunat MS "
        psSql = psSql & " LEFT Join MovProvisionAgencia MPA on MPA.nmovnro=ms.nmovnro "
        psSql = psSql & " Where cmovnro Like '" & Me.txtFFiltro.Text & "%' and bvigente=1 AND cAgeCod is NULL"
      End If
      
      Set CargaObjeto = oCon.CargaRecordSet(psSql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaObjetoErr:
'   Call RaiseError(MyUnhandledError, "DObjeto:CargaObjeto Method")
    MsgBox Err.Description
End Function


Private Function CargaDetalleDocumento(ByVal psMovNro As String, Optional psTipo As String = "") As ADODB.Recordset
    Dim oCon As New DConecta
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim sCTAS As String
    
    sql = " select cCtaContCod from OPECTA where copecod = '" & gsOpeCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    If Not RSVacio(rs) Then
        While Not rs.EOF
            If sCTAS = "" Then
                sCTAS = "'" & rs!cCtaContCod & "'"
            Else
                sCTAS = sCTAS & ",'" & rs!cCtaContCod & "'"
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    If psTipo = "" Then
        sql = " SELECT distinct md.dDocFecha, doc.cDocAbrev, md.nDocTpo, md.cDocNro, Prov.cpersnombre cPersona, Prov.nPersPersoneria, M.cMovDesc, Prov.cPersCod,ps.nImporteCoactivo,ps.ntpocambio ,"
        sql = sql & "       m.cMovNro, m.nMovNro, mc.cCtaContCod, ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1 as nMovImporte,( mc.nMovImporte) * -1 as nMovImporteSoles,Provi.cPersIDnro, cMovNroRes,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,''),0) MontoPagadoSUNAT,ref.cMovNro MovPago,isnull(dbo.GetMontoPagadoSUNAT(m.nmovnro,1),0) MontoPagadoSUNATsoles ,ps.ntpocambio ,0 Penalidad  ,'' MovEnvio , '' Agencia"
    Else
        sql = " SELECT distinct Prov.cpersnombre cPersona,"
        sql = sql & "        Prov.cPersCod,  sum(ps.nImportecoactivo ) as nImportecoactivo "
    End If
    sql = sql & " FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro "
    sql = sql & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
    sql = sql & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro "
    sql = sql & "              "
    sql = sql & "              "
    sql = sql & " LEFT JOIN (SELECT mr.nMovNro,m1.cmovnro ,mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
    sql = sql & "                   WHERE m1.copecod not in ('" & OpeCGOpeProvPagoPagUNAT & "','401581') and  m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
    sql = sql & "                  ) ref ON  ref.nMovNroRef = m.nMovNro "
    sql = sql & "             JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod "
    sql = sql & "             left join Persid ProvI  ON ProvI.cPersCod = mg.cPersCod and cPersIdTpo=2"
    sql = sql & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo "
    sql = sql & "          join movControlPagoSunat ps ON ps.nmovnro = m.nMovNro and ps.cmovnro='" & psMovNro & "'"
    sql = sql & " WHERE  m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
    'sql = sql & " and mc.cCtaContCod in(" & sCTAS & ") and ps.bVigente=1 and bvalido=1 AND ISNULL(me.nMovMeImporte,mc.nMovImporte)<0 AND ISNULL(Mc.nTipoPago,0)=0"
    sql = sql & " and mc.cCtaContCod in(" & sCTAS & ") and ps.bVigente=1 and bvalido=1 AND ISNULL(Mc.nTipoPago,0)=0" 'EJVG20140430
    If psTipo <> "" Then
        sql = sql & " group by Prov.cpersnombre , Prov.cPersCod "
    End If

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set CargaDetalleDocumento = rs
End Function

Private Sub optAge_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.txtAge.Text = ""
            Me.lblAgencia.Caption = ""
            Me.txtAge.Enabled = False
            lsTipoB = "T"
        Case 1
            Me.txtAge.Enabled = True
            lsTipoB = "A"
        Case 2
            Me.txtAge.Text = ""
            Me.lblAgencia.Caption = ""
            Me.txtAge.Enabled = False
            lsTipoB = "L"
    End Select
End Sub

Private Sub txtAge_EmiteDatos()
    Me.lblAgencia = Me.txtAge.psDescripcion
End Sub

Private Sub txtFFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtProveedor_EmiteDatos()
    If Me.txtProveedor <> "" Then
        Me.cmdProcesar.Enabled = False
        Me.cmdImprimir.Enabled = False
        DoEvents
        MuestraDetalle
        Me.cmdProcesar.Enabled = True
        Me.cmdImprimir.Enabled = True
    End If
End Sub

Private Sub MuestraDetalle()
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    Me.fg.Clear
    Me.fg.Rows = 2
    Me.fg.FormaCabecera
    Set rs = CargaDetalleDocumento(Me.txtProveedor.Text)
    If Not RSVacio(rs) Then
        Me.cmdImprimir.Enabled = True
        While Not rs.EOF
            fg.AdicionaFila
            nItem = fg.row
            fg.TextMatrix(nItem, 1) = nItem
            fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
            fg.TextMatrix(nItem, 4) = rs!dDocFecha
            fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPErsona, True)
            fg.TextMatrix(nItem, 6) = rs!cMovDesc
            fg.TextMatrix(nItem, 7) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 8) = rs!cPersCod
            fg.TextMatrix(nItem, 9) = rs!cMovNro
            fg.TextMatrix(nItem, 10) = rs!nMovNro
            fg.TextMatrix(nItem, 11) = rs!nDocTpo
            fg.TextMatrix(nItem, 12) = rs!cDocNro
            fg.TextMatrix(nItem, 13) = rs!cCtaContCod
            fg.TextMatrix(nItem, 14) = GetFechaMov(rs!cMovNro, True)
            fg.TextMatrix(nItem, 24) = rs!Penalidad
            fg.TextMatrix(nItem, 25) = rs!Movenvio
            fg.TextMatrix(nItem, 26) = rs!Agencia
            
            If rs!nMovImporte = rs!nMovImporteSoles Then
                fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                fg.TextMatrix(nItem, 23) = "SOLES"
            Else
                fg.TextMatrix(nItem, 15) = Format(Round((rs!nMovImporte - rs!MontoPagadoSUNAT) * rs!nTpoCambio, 2), gsFormatoNumeroView)
                fg.TextMatrix(nItem, 23) = "DOLARES"
            End If
            
            fg.TextMatrix(nItem, 16) = IIf(IsNull(rs!cPersIDnro), "RUC NO REGISTRADO", Trim(rs!cPersIDnro))
            fg.TextMatrix(nItem, 17) = Format(IIf(IsNull(rs!nimportecoactivo), "0.00", rs!nimportecoactivo), gsFormatoNumeroView)
            fg.TextMatrix(nItem, 18) = IIf(CCur(fg.TextMatrix(nItem, 7)) = CCur(fg.TextMatrix(nItem, 15)) + CCur(rs!MontoPagadoSUNAT), "SOLES", "DOLARES")
            fg.TextMatrix(nItem, 19) = IIf(IsNull(rs!cMovNroRes), "", rs!cMovNroRes)
            fg.TextMatrix(nItem, 20) = IIf(IsNull(rs!MovPago), "", rs!MovPago)
            fg.TextMatrix(nItem, 21) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 22) = Format(rs!MontoPagadoSUNATsoles, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 2) = IIf(IsNull(rs!nimportecoactivo), "0", "1")
            rs.MoveNext
        Wend
    Else
        Me.cmdImprimir.Enabled = False
    End If
End Sub


Public Sub GenerarExcel(Optional psTipo As String = "")
    Dim fs              As Scripting.FileSystemObject
    Dim xlAplicacion    As Excel.Application
    Dim xlLibro         As Excel.Workbook
    Dim xlHoja1         As Excel.Worksheet
    Dim lbExisteHoja    As Boolean
    Dim liLineas        As Integer
    Dim i               As Integer
    Dim glsArchivo      As String
    Dim lsNomHoja       As String
    Dim rs              As New ADODB.Recordset
    Dim rs1             As New ADODB.Recordset
    Dim sCadenaA        As String
    Dim sCadenaB        As String
    

    glsArchivo = "Reporte_ProveedorxArchivo" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLSX"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

    lbExisteHoja = False
    lsNomHoja = "Proveedores"
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

    xlAplicacion.Range("A1:A1").ColumnWidth = 2
    xlAplicacion.Range("B1:B1").ColumnWidth = 6
    xlAplicacion.Range("c1:c1").ColumnWidth = 30
    xlAplicacion.Range("d1:d1").ColumnWidth = 11
    xlAplicacion.Range("e1:e1").ColumnWidth = 40
    xlAplicacion.Range("f1:f1").ColumnWidth = 10.3
    xlAplicacion.Range("g1:g1").ColumnWidth = 11.3
    xlAplicacion.Range("h1:h1").ColumnWidth = 12.3
    xlAplicacion.Range("j1:j1").ColumnWidth = 19
    xlAplicacion.Range("l1:l1").ColumnWidth = 20.5
    xlAplicacion.Range("A1:Z100").Font.Size = 9

    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(2, 2) = "REPORTE DE PROVEEDORES X ARCHIVO MN_ME " & IIf(psTipo <> "", "RESUMEN", "DETALLE")
    xlHoja1.Cells(3, 2) = "INFORMACION  AL  " & Format(gdFecSis, "dd/mm/yyyy")
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 10)).Merge True
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).Merge True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 10)).HorizontalAlignment = xlCenter
    
    If psTipo = "" Then
        xlHoja1.Cells(5, 2) = "Nro"
        xlHoja1.Cells(5, 3) = "Comprobante"
        xlHoja1.Cells(5, 4) = "Emision"
        xlHoja1.Cells(5, 5) = "Proveedor"
        xlHoja1.Cells(5, 6) = "Importe"
        xlHoja1.Cells(5, 7) = "Valor Soles"
        xlHoja1.Cells(5, 8) = "Retencion S/."
        xlHoja1.Cells(5, 9) = "Moneda"
        xlHoja1.Cells(5, 10) = "Monto Pagado SUNAT"
        xlHoja1.Cells(5, 11) = "Tipo Cambio"
        xlHoja1.Cells(5, 12) = "Monto Pagado SUNAT S/."
    Else
        xlHoja1.Cells(5, 2) = "Nro"
        xlHoja1.Cells(5, 3) = "Proveedor"
        xlHoja1.Cells(5, 4) = "Retencion S/."
    End If
    liLineas = 5
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, IIf(psTipo = "", 12, 4))).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, IIf(psTipo = "", 12, 4))).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, IIf(psTipo = "", 12, 4))).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, IIf(psTipo = "", 12, 4))).Interior.Color = RGB(159, 206, 238)
    
    Set rs = CargaDetalleDocumento(Me.txtProveedor.Text, psTipo)
    If Not RSVacio(rs) Then
        While Not rs.EOF
            liLineas = liLineas + 1
            xlHoja1.Cells(liLineas, 2) = rs.Bookmark
            If psTipo = "" Then
                xlHoja1.Cells(liLineas, 3) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
                xlHoja1.Cells(liLineas, 4) = "'" & rs!dDocFecha
                xlHoja1.Cells(liLineas, 5) = PstaNombre(rs!cPErsona, True)
                xlHoja1.Cells(liLineas, 6) = Format(rs!nMovImporte, gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 7) = Format((rs!nMovImporte * rs!nTpoCambio) - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 8) = Format(IIf(IsNull(rs!nimportecoactivo), "0.00", rs!nimportecoactivo), gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 9) = IIf(CCur(rs!nMovImporte) = CCur(rs!nMovImporteSoles), "SOLES", "DOLARES")
                xlHoja1.Cells(liLineas, 10) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                xlHoja1.Cells(liLineas, 11) = rs!nTpoCambio
                xlHoja1.Cells(liLineas, 12) = Format(rs!MontoPagadoSUNATsoles, gsFormatoNumeroView)
            Else
                xlHoja1.Cells(liLineas, 3) = PstaNombre(rs!cPErsona, True)
                xlHoja1.Cells(liLineas, 4) = Format(rs!nimportecoactivo, gsFormatoNumeroView)
                
            End If
            rs.MoveNext
        Wend
    End If

    xlHoja1.Range("D:D").Style = "comma"
    xlHoja1.Range("h:l").Style = "comma"

    
    ExcelCuadro xlHoja1, 2, 5, IIf(psTipo = "", 12, 4), liLineas
    
    xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo

    MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Imprimio Excel "
                Set objPista = Nothing
                '****
End Sub
