VERSION 5.00
Begin VB.Form frmCajaGenMantCtasVinculados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Adeudados Vinculados"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEntidad 
      Caption         =   "Adeudos"
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
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1875
         _ExtentX        =   3307
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
         sTitulo         =   ""
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2010
         TabIndex        =   3
         Top             =   210
         Width           =   5235
      End
   End
   Begin VB.Frame FraGenerales 
      Caption         =   "Adeudados Vinculados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8940
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   1110
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1110
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1110
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2280
         Width           =   1080
      End
      Begin Sicmact.FlexEdit fgAdeudadoVin 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3201
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Cuenta IF-Nro Cuenta-Institucion-Monto Prestado-Saldo Capital-Valor"
         EncabezadosAnchos=   "500-2000-2500-3500-1800-1800-0"
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-1-2-2-3"
         CantDecimales   =   4
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCajaGenMantCtasVinculados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCtaIf As NCajaCtaIF
Dim gNroFilas As Integer

Dim sPersCodV As String
Dim nIFTpoV As String
Dim sCtaIFCodV As String

Private Sub cmdAgregar_Click()
    fgAdeudadoVin.AdicionaFila
   
    cmdGrabar.Enabled = False
    cmdSalir.Enabled = False
    
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    CmdCancelar.Enabled = True
        
    fgAdeudadoVin.Col = 2
    Call fgAdeudadoVin_RowColChange
    fgAdeudadoVin.Col = 1
       
    
End Sub

Private Sub cmdCancelar_Click()
If fgAdeudadoVin.Row <= gNroFilas Then
   Exit Sub
End If
fgAdeudadoVin.EliminaFila (fgAdeudadoVin.Row)
fgAdeudadoVin_RowColChange
'gFilaActual = fgPlazoInt.Rows
If fgAdeudadoVin.TextMatrix(1, 1) = "" Then
   cmdGrabar.Enabled = False
End If

cmdAgregar.Enabled = True
cmdEliminar.Enabled = True
CmdCancelar.Enabled = False
cmdSalir.Enabled = True

End Sub

Private Sub cmdEliminar_Click()
Dim I          As Integer
Dim oCtaIf     As NCajaCtaIF
Dim oMov       As New DMov
Dim gsMovNro   As String
Dim sPersCod As String
Dim sIFTpo As String
Dim sCtaIFCod As String

Set oCtaIf = New NCajaCtaIF
Set oMov = New DMov

    If fgAdeudadoVin.TextMatrix(1, 2) = "" Then
        MsgBox "No existe elemento para Eliminar", vbInformation, "Aviso"
        Me.fgAdeudadoVin.EliminaFila (Me.fgAdeudadoVin.Row)
        Exit Sub
    End If

        sPersCod = Mid(fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 1), 4, 13)
        sIFTpo = Mid(fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 1), 1, 2)
        sCtaIFCod = Mid(fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 1), 18, 10)

   'gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   If MsgBox(" ¿ Esta seguro de eliminar esta garantia ? ", vbQuestion + vbYesNo) = vbYes Then
       If oCtaIf.EliminarAdeudoVinculado(sPersCod, sIFTpo, sCtaIFCod) = False Then
          MsgBox " Adeudo no fue Eliminado ", vbInformation, "Aviso"
          Exit Sub
       Else
          Me.fgAdeudadoVin.EliminaFila (Me.fgAdeudadoVin.Row)
          Me.cmdGrabar.Enabled = False
          sPersCodV = Mid(txtBuscaEntidad.Text, 4, 13)
          nIFTpoV = Mid(txtBuscaEntidad.Text, 1, 2)
          sCtaIFCodV = Mid(txtBuscaEntidad.Text, 18, 10)
          Call CargaAdeudos(sPersCodV, nIFTpoV, sCtaIFCodV)
          
       End If
   End If
If fgAdeudadoVin.TextMatrix(1, 1) = "" Then
    cmdGrabar.Enabled = False
End If
End Sub

Private Sub CmdGrabar_Click()
Dim oCtaIf As NCajaCtaIF
Dim lsMovNro As String
Dim oMov As DMov

Dim sPersCod As String
Dim nIFTpo As String
Dim sCtaIFCod As String
Dim nImporte As Currency
Dim I As Integer
Dim nValor As Integer

Set oMov = New DMov
Set oCtaIf = New NCajaCtaIF
    If MsgBox("Desea Grabar las Cuentas Vinculadas??", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
        '------Institucion Vinculada
        sPersCodV = Mid(Trim(txtBuscaEntidad.Text), 4, 13)
        nIFTpoV = Mid(txtBuscaEntidad.Text, 1, 2)
        sCtaIFCodV = Mid(txtBuscaEntidad.Text, 18, 10)
         
                        
       nValor = oCtaIf.GrabaAdeudoVinculado(sPersCodV, nIFTpoV, sCtaIFCodV, gnTipoCambioEuro, fgAdeudadoVin.GetRsNew)
       If nValor = 1 Then
           MsgBox "Datos Grabados Satisfactoriamente", vbInformation, "Aviso"
        Else
           MsgBox "No se pudo Grabar", vbCritical, "Aviso"
           Exit Sub
       End If
     'lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     'oMov.InsertaMov lsMovNro, gsOpeCod, "Adeudos Vinculados", gMovEstContabNoContable, gMovFlagVigente
     sPersCodV = Mid(txtBuscaEntidad.Text, 4, 13)
     nIFTpoV = Mid(txtBuscaEntidad.Text, 1, 2)
     sCtaIFCodV = Mid(txtBuscaEntidad.Text, 18, 10)

    CargaAdeudos sPersCodV, nIFTpoV, sCtaIFCodV
         
Set oCtaIf = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgAdeudadoVin_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim sPersCod As String
Dim nIFTpo As String
Dim sCtaIFCod As String

Set oCtaIf = New NCajaCtaIF
 If fgAdeudadoVin.TextMatrix(pnRow, 1) = "" And psDataCod <> "" Then
       'if fgAdeudadoVin.te
       MsgBox "Cuenta ya Asignada", vbInformation, "Aviso"
       fgAdeudadoVin.Rows = fgAdeudadoVin.Rows - 1
       cmdAgregar.Enabled = True
       cmdEliminar.Enabled = True
       CmdCancelar.Enabled = False
       cmdSalir.Enabled = True
       Exit Sub
    End If

    sPersCod = Mid(fgAdeudadoVin.TextMatrix(pnRow, 1), 4, 13)
    nIFTpo = Mid(fgAdeudadoVin.TextMatrix(pnRow, 1), 1, 2)
    sCtaIFCod = Mid(fgAdeudadoVin.TextMatrix(pnRow, 1), 18, 10)
    
Set rs1 = oCtaIf.GetCtaIFMontos(sPersCod, nIFTpo, sCtaIFCod)
If Not rs1.EOF Then
        
    fgAdeudadoVin.TextMatrix(pnRow, 3) = oCtaIf.NombreIF(sPersCod)
    fgAdeudadoVin.TextMatrix(pnRow, 4) = rs1(0)
    fgAdeudadoVin.TextMatrix(pnRow, 5) = rs1(1)
    
    cmdGrabar.Enabled = True
    cmdSalir.Enabled = True
End If
Set oCtaIf = Nothing

End Sub

Private Sub fgAdeudadoVin_RowColChange()
Dim I As Integer
Dim nFilaAct As Integer

With fgAdeudadoVin
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(.Row, 6)) = 0 Then
            nFilaAct = .Row
            Exit For
        End If
    Next
End With
If nFilaAct = 0 Then
    fgAdeudadoVin.lbEditarFlex = False
Else
    fgAdeudadoVin.lbEditarFlex = True
    If Val(fgAdeudadoVin.TextMatrix(fgAdeudadoVin.Row, 6)) = 1 Then
        fgAdeudadoVin.Row = nFilaAct
    End If
End If
End Sub

Private Sub Form_Load()
Dim oOpe As New DOperacion
Dim sSql As String
Dim oCon As New DConecta
Dim rs As New ADODB.Recordset


Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
'txtBuscarCtaIF.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
Set oOpe = Nothing

txtBuscaEntidad.psRaiz = "Cuentas de Instituciones Financieras"
Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")


sSql = " SELECT   CASE WHEN NIVEL =1 THEN CPERSCOD ELSE CPERSCOD + '.' + cCtaIFCod END AS CODIGO ,"
sSql = sSql & " Convert(char(40),CTAIFDESC)  as CTAIFDESC, Nivel"
sSql = sSql & " FROM ("
sSql = sSql & "     SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod,"
sSql = sSql & "     CONVERT(CHAR(40),ISNULL( (SELECT LEFT(cDescripcion,22)"
sSql = sSql & "     from ctaifadeudados cia"
sSql = sSql & "     join coloclineacredito cl on cl.cLineaCred = cia.cCodLinCred"
sSql = sSql & " Where cia.cPersCod = CI.cPersCod And cia.cIFTpo = CI.cIFTpo And cia.CCTAIFCOD = CI.CCTAIFCOD"
sSql = sSql & "     ) + ' ','') + CI.cCtaIFDesc) AS CTAIFDESC,                  LEN(CI.cCtaIFCod) AS Nivel,"
sSql = sSql & " i.cIFTpo , i.bCanje"
sSql = sSql & " FROM    INSTITUCIONFINANC I"
sSql = sSql & " JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo"
sSql = sSql & " WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X')  AND SUBSTRING(CI.cCtaIfCod,3,1)like '[12]'"
sSql = sSql & " AND CI.cIFTpo+CI.cCtaIfCod LIKE '_[1234567]_[5]%'"
sSql = sSql & " and (LEN(ci.cCtaIFCod) > 3"
sSql = sSql & " or EXISTS(  Select cIFTpo FROM CtaIF civ WHERE"
sSql = sSql & " civ.cIFTpo = CI.cIFTpo And civ.cPersCod = CI.cPersCod And civ.CCTAIFCOD"
sSql = sSql & "         like ci.cCtaIFCod + '_%' )"
sSql = sSql & "     )"
sSql = sSql & " Union"
sSql = sSql & " SELECT  I.cIFTpo + '.' + I.CPERSCOD as CPERSCOD, '' AS CTAIF, P.CPERSNOMBRE , 1 AS NIVEL ,"
sSql = sSql & " i.cIFTpo , i.bCanje"
sSql = sSql & " FROM    INSTITUCIONFINANC I"
sSql = sSql & " JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD"
sSql = sSql & " JOIN ("
sSql = sSql & "     SELECT  CI.cIFTpo, CI.CPERSCOD"
sSql = sSql & "     FROM    CTAIF CI"
sSql = sSql & "     WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X')  AND SUBSTRING(CI.cCtaIfCod,3,1)='1'"
sSql = sSql & "     AND CI.cIFTpo+CI.cCtaIfCod LIKE '_[1234567]_[5]%'"
sSql = sSql & "     ) AS C1                  ON  C1.cIFTpo=I.cIFTpo AND C1.CPERSCOD= I.CPERSCOD"
sSql = sSql & " Union"
sSql = sSql & " Select  Replace(Str(nConsValor,2,0),' ','0') as cPerscod, '' as CtaIf ,"
sSql = sSql & "     cConsDescripcion , 0 AS  NIVEL, Replace(Str(nConsValor,2,0),' ','0') as cIFTpo,"
sSql = sSql & "     0 as bCanje"
sSql = sSql & " From Constante"
sSql = sSql & " Where nConsCod Like 4001 And nConsValor <> 4001"
sSql = sSql & " AND  Replace(Str(nConsValor,2,0),' ','0') IN ("
sSql = sSql & "         SELECT  DISTINCT I.cIFTpo"
sSql = sSql & "         FROM    INSTITUCIONFINANC I"
sSql = sSql & "         JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo"
sSql = sSql & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X')"
sSql = sSql & "         AND SUBSTRING(CI.cCtaIfCod,3,1)like '[12]'  AND CI.cIFTpo+CI.cCtaIfCod LIKE '_[1234567]_[5]%' )"
sSql = sSql & " ) AS CTASIF"
sSql = sSql & " Where Nivel <= 7  ORDER BY CPERSCOD, cCtaIFCod"
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
fgAdeudadoVin.rsTextBuscar = rs
oCon.CierraConexion
Set oOpe = Nothing

End Sub


Private Sub txtBuscaEntidad_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Dim sPersCodV As String
Dim nIFTpoV As String
Dim sCtaIFCodV As String


Set oCtaIf = New NCajaCtaIF
lblCtaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
If txtBuscaEntidad <> "" Then
        cmdSalir.Enabled = True
        cmdGrabar.Enabled = True
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        CmdCancelar.Enabled = False
 End If
Set oCtaIf = Nothing
    sPersCodV = Mid(txtBuscaEntidad.Text, 4, 13)
    nIFTpoV = Mid(txtBuscaEntidad.Text, 1, 2)
    sCtaIFCodV = Mid(txtBuscaEntidad.Text, 18, 10)

If txtBuscaEntidad.Text <> "" Then
    CargaAdeudos sPersCodV, nIFTpoV, sCtaIFCodV
End If
'fgAdeudadoVin.rsFlex = GetCtaIFVinculados(sPersCodV, nIfTpoV, sCtaIFCodV)

End Sub
Public Sub CargaAdeudos(ByVal sPersCodV As String, ByVal nIFTpoV As String, ByVal sCtaIFCodV As String)
Dim res As New ADODB.Recordset
Dim oCtaIf As NCajaCtaIF
    
    Set res = New ADODB.Recordset
    
    Me.fgAdeudadoVin.Clear
    Me.fgAdeudadoVin.Rows = 2
    Me.fgAdeudadoVin.FormaCabecera
    'gFilaActual = 0
    Set oCtaIf = New NCajaCtaIF
    Set res = GetCtaIFVinculados(sPersCodV, nIFTpoV, sCtaIFCodV)
    If RSVacio(res) Then
        Me.fgAdeudadoVin.Rows = 2
    Else
        While Not res.EOF
            Me.fgAdeudadoVin.AdicionaFila
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 1) = res!Codigo
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 2) = res!cCtaIFDesc
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 3) = res!cPersNombre
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 4) = res!nMontoPrestado
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 5) = res!nSaldoCap
            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 6) = 1

'            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 7) = res!bEuros
'            Me.fgAdeudadoVin.TextMatrix(Me.fgAdeudadoVin.Row, 8) = res!nMontoEuros

            res.MoveNext
        Wend
        Me.fgAdeudadoVin.lbEditarFlex = True
        'Call fgPlazoInt_RowColChange
    End If
    gNroFilas = res.RecordCount
If res.RecordCount >= 1 Then
        cmdGrabar.Enabled = True
        cmdSalir.Enabled = True
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        CmdCancelar.Enabled = False
    Else
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        CmdCancelar.Enabled = True
        cmdGrabar.Enabled = True
    End If

'fgAdeudadoVin.rsFlex = GetCtaIFVinculados(sPersCodV, nIfTpoV, sCtaIFCodV)

End Sub
