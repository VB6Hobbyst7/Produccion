VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpeRegVentanilla 
   Caption         =   "Operaciones con Ventanilla"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   Icon            =   "frmOpeRegVentanilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraAge 
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
      Height          =   615
      Left            =   90
      TabIndex        =   18
      Top             =   2460
      Width           =   7560
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   345
         Left            =   1065
         TabIndex        =   2
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   -2147483647
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2445
         TabIndex        =   20
         Top             =   210
         Width           =   4965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
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
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importe "
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
      Height          =   780
      Left            =   90
      TabIndex        =   15
      Top             =   6150
      Width           =   3000
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   285
         Width           =   1905
      End
      Begin VB.Label lblSimbolo 
         Alignment       =   1  'Right Justify
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame fraCarta 
      Height          =   600
      Left            =   5415
      TabIndex        =   13
      Top             =   0
      Width           =   2235
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6090
      TabIndex        =   8
      Top             =   6360
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   4650
      TabIndex        =   7
      Top             =   6360
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción"
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
      Height          =   975
      Left            =   90
      TabIndex        =   11
      Top             =   5130
      Width           =   7575
      Begin VB.TextBox txtMovDesc 
         Height          =   600
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7320
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Concepto"
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
      Height          =   1785
      Left            =   90
      TabIndex        =   10
      Top             =   3300
      Width           =   7575
      Begin Sicmact.TxtBuscar txtCtaCod 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
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
      Begin VB.TextBox txtCtaDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1740
         TabIndex        =   4
         Top             =   240
         Width           =   5685
      End
      Begin Sicmact.FlexEdit fgObj 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   1720
         Cols0           =   8
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
         EncabezadosAnchos=   "350-600-1500-3500-0-1000-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-C-C-C"
         FormatosEdit    =   "0-0-3-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraVentanilla 
      Caption         =   "Pagos realizados"
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
      Height          =   1800
      Left            =   90
      TabIndex        =   9
      Top             =   585
      Width           =   7560
      Begin MSComctlLib.ListView lstPago 
         Height          =   1395
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Persona"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nro.Doc."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CodPers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "cGlosa"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   6270
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   609
      _Version        =   393217
      TextRTF         =   $"frmOpeRegVentanilla.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOpeRegVentanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs   As ADODB.Recordset
Dim sSql As String
Dim sCtaPend As String
Dim oContFunct As NContFunciones

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Me.lstPago.ListItems.Count = 0 Then
   MsgBox "No se detectaron Ingresos por Ventanilla", vbInformation, "¡Aviso!"
   Exit Function
End If
If txtCtaCod = "" Then
   MsgBox "Falta indicar Concepto motivo del Ingreso", vbInformation, "¡Aviso!"
   Exit Function
End If
If txtMovDesc = "" Then
   MsgBox "Falta indicar descripción de la Operación", vbInformation, "¡Aviso!"
   Exit Function
End If
If Me.txtAgeCod = "" Then
    MsgBox " Seleccione la agencia donde se realizo la operacion ", vbInformation, "¡Aviso!"
    Exit Function
End If
If Not ValidaFechaContab(Me.txtFecha, gdFecSis) Then
   txtFecha.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim I       As Integer
Dim sImpre  As String
Dim oMov As DMov
Dim oFun As New NContFunciones
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnItem   As Long
Dim lsSubCta As String
Dim lnOrdenObj As Long

If Not ValidaDatos() Then
   Exit Sub
End If

If MsgBox(" ¿ Seguro que desea Grabar datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
Set oMov = New DMov
Set oFun = New NContFunciones

gsMovNro = oMov.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
If Not oFun.PermiteModificarAsiento(gsMovNro, False) Then
   MsgBox "Imposible registrar Regularización con fecha de Mes ya Cerrado", vbInformation, "!Aviso¡"
   oMov.RollbackTrans
   Exit Sub
End If

oMov.BeginTrans
lnItem = 1
lbTrans = True
gsGlosa = txtMovDesc
gnImporte = nVal(lstPago.SelectedItem.SubItems(2))

oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(gsMovNro)
oMov.InsertaMovCont lnMovNro, gnImporte, 0, ""
oMov.InsertaMovCta lnMovNro, lnItem, sCtaPend, gnImporte

Dim rsObjetos As ADODB.Recordset
Set rsObjetos = fgObj.GetRsNew()
If Not rsObjetos Is Nothing Then
    rsObjetos.MoveFirst
    Do While Not rsObjetos.EOF
        If rsObjetos!CTACONT = txtCtaCod.Text Then
            lsSubCta = lsSubCta + rsObjetos![SubCta]
        End If
        rsObjetos.MoveNext
    Loop
End If
lnItem = lnItem + 1
oMov.InsertaMovCta lnMovNro, lnItem, txtCtaCod & lsSubCta, gnImporte * -1


If Not rsObjetos Is Nothing Then
    rsObjetos.MoveFirst
    Do While Not rsObjetos.EOF
        If rsObjetos!CTACONT = txtCtaCod And rsObjetos!ItemCtaCont = 1 Then
            lnOrdenObj = rsObjetos!Ord
            Select Case rsObjetos!ObjPadre
               Case ObjCMACAgencias
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código, ""
               Case ObjCMACArea
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, "", rsObjetos!Código
               Case ObjCMACAgenciaArea
                    Dim lsAgeCod As String
                    Dim lsAreaCod As String
                    lsAgeCod = Mid(rsObjetos!Código, 4, 2)
                    lsAreaCod = Mid(rsObjetos!Código, 1, 3)
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                    oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, lsAgeCod, lsAreaCod
               Case ObjDescomEfectivo
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                    oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
               Case ObjEntidadesFinancieras
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                    oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrdenObj, Mid(rsObjetos!Código, 4, 13), Left(rsObjetos!Código, 2), Mid(rsObjetos!Código, 18)
               Case Else
                    oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
            End Select
        End If
        rsObjetos.MoveNext
    Loop
End If
lnOrdenObj = lnOrdenObj + 1

'La persona del Concepto seleccionado
oMov.InsertaMovGasto lnMovNro, lstPago.SelectedItem.SubItems(4), ""
oMov.InsertaMovRef lnMovNro, lstPago.SelectedItem.SubItems(3), txtAgeCod
If Mid(gsOpeCod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, gsMovNro
End If
oMov.ActualizaSaldoMovimiento gsMovNro, "+"
oMov.CommitTrans
'oMov.RollbackTrans
Dim oCon As New DConecta
If Not gbBitCentral Then
   oCon.AbreConexion 'Remota txtAgeCod, False, False, "01"
   sSql = "UPDATE TransRef SET cCodAge = '" & gsCodCMAC & txtAgeCod & "', cNroTran = '" & gsMovNro & "' WHERE cNroTranRef = '" & lstPago.SelectedItem.SubItems(3) & "' "
   oCon.Ejecutar sSql
   oCon.CierraConexion
   Set oCon = Nothing
End If
glAceptar = True
Dim oImp As NContImprimir
Set oImp = New NContImprimir
oImp.Inicio gsNomCmac, gsNomAge, Format(gdFecSis, gsFormatoFechaView)
sImpre = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "ASIENTO DE REGULARIZACION DE INGRESO")
sImpre = oImp.ImprimeRecibo(gsMovNro, False, "C O M P R O B A N T E   D E   R E G U L A R I Z A C I O N") & oImpresora.gPrnSaltoPagina & sImpre & oImpresora.gPrnSaltoPagina & sImpre
Set oImp = Nothing
EnviaPrevio sImpre, gsOpeDesc, gnLinPage, False

If MsgBox(" ¿ Desea registrar otra operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Unload Me
   Exit Sub
End If
LimpiaFlex
lstPago.ListItems.Remove lstPago.SelectedItem.Index
If lstPago.ListItems.Count = 0 Then
   MsgBox "No existen más Ingresos para Regularizar", vbInformation, "Aviso!"
   Unload Me
   Exit Sub
End If
txtCtaCod = ""
txtCtaDes = ""
txtMovDesc = ""
lstPago.SetFocus
End Sub

'Private Sub AsignaObjetos(sCtaCod As String, sCtaDes As String)
'Dim rsObj As New ADODB.Recordset
'Dim nNiv As Integer
'Dim nObj As Integer
'Dim nObjs As Integer
'   LimpiaFlex
'      txtCtaCod = sCtaCod
'      txtCtaDes = sCtaDes
'      sSql = "SELECT MAX(cCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
'      Set rs = CargaRecord(sSql)
'      nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
'      For nObj = 1 To nObjs
'         sSql = "SELECT co.cObjetoCod, co.cCtaObjOrden, o.nObjetoNiv, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre FROM CtaObj co JOIN Objeto o ON o.cObjetoCod = co.cObjetoCod WHERE co.cCtaContCod = '" & sCtaCod & "' and co.cCtaObjOrden = '" & nObj & "'"
'         Set rs = CargaRecord(sSql)
'         AceptaOK = False
'         If rs!cObjetoCod <> "00" Then
'            AdicionaObj sCtaCod, "001", rs
'            nNiv = rs!nObjetoNiv + rs!nCtaObjNiv
'            If rs.RecordCount = 1 Then 'Una Cuenta por Orden
'               sSql = " spGetTreeObj '" & rs!cObjetoCod & "', " & nNiv & ",'" & rs!cCtaObjFiltro & "'"
'            Else  'Para varios Objetos en una misma Orden asignado a la Cuenta
'               sSql = "SELECT o.* FROM Objeto O JOIN CtaObj co ON co.cObjetoCod = substring(o.cObjetoCod,1,LEN(co.cObjetoCod)) WHERE co.cCtaContCod = '" & sCtaCod & "' and co.cCtaObjOrden = '" & nObj & "' and Exists (select * from Objeto as a where a.cobjetocod like o.cobjetocod+'%' and a.nObjetoNiv = " & nNiv & ") order by o.cobjetocod "
'            End If
'            If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
'            Set rsObj = dbCmact.Execute(sSql)
'            If Not rsObj.EOF Then
'               frmDescObjeto.Inicio rsObj, "", nNiv
'               If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
'               If frmDescObjeto.lOk Then
'                  AceptaOK = True
'               End If
'            End If
'         End If
'         If AceptaOK Then
'            fgObj.TextMatrix(fgObj.Row, 2) = gaObj(0, 0, UBound(gaObj, 3) - 1) 'CodObj
'            fgObj.TextMatrix(fgObj.Row, 3) = gaObj(0, 1, UBound(gaObj, 3) - 1) 'DesObj
'            sSql = "SELECT cCtaObjMascara FROM CtaObjFiltro WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod = '" & fgObj.TextMatrix(fgObj.Row, 2) & "'"
'            If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
'            rsObj.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'            If rsObj.EOF Then
'               fgObj.TextMatrix(fgObj.Row, 4) = ""
'            Else
'               fgObj.TextMatrix(fgObj.Row, 4) = rsObj!cCtaObjMascara
'            End If
'         Else
'            If rs!cObjetoCod <> "00" Then
'               LimpiaFlex
'               txtCtaCod = ""
'               txtCtaDes = ""
'               RSClose rs
'               Exit Sub
'            End If
'         End If
'         RSClose rs
'      Next
'End Sub
'
Private Sub LimpiaFlex()
fgObj.Clear
fgObj.Rows = 2
fgObj.FormaCabecera
End Sub

Private Sub cmdSalir_Click()
glAceptar = False
Unload Me
End Sub

Private Function CargaLista(psCodAge As String)
Dim oNeg     As NNegOpePendientes
Dim lvItem   As ListItem
Dim lsOpeCod As String
On Error GoTo ErrCargaLista
Set oNeg = New NNegOpePendientes
If gbBitCentral Then
   lsOpeCod = gOtrOpeIngresosoCajaGeneral
Else
   lsOpeCod = gCapIngresoRegulaCG
End If
Set rs = oNeg.CargaIngresoVentanillaPendiente(lsOpeCod, gsCodCMAC & psCodAge, Mid(gsOpeCod, 3, 1))
lstPago.ListItems.Clear
Do While Not rs.EOF
   Set lvItem = lstPago.ListItems.Add()
   lvItem.Text = Format(rs!DFECTRAN, gsFormatoFechaView)
   lvItem.SubItems(1) = PstaNombre(rs!cNomPers)
   lvItem.SubItems(2) = Format(rs!nMonTran, gsFormatoNumeroView)
   lvItem.SubItems(3) = rs!nMovNro
   If gbBitCentral Then
      lvItem.SubItems(4) = rs!cPersCodIF
   Else
      lvItem.SubItems(4) = gsCodCMAC & rs!cPersCodIF
   End If
   lvItem.SubItems(5) = Replace(rs!cGlosa, Chr(13), "")
   rs.MoveNext
Loop
RSClose rs
Set oNeg = Nothing
Exit Function
ErrCargaLista:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Sub Form_Load()
Dim lvItem As ListItem
Dim oCaja As DCajaGeneral
CentraForm Me
Me.Caption = gsOpeDesc
Set oContFunct = New NContFunciones

Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeCta(gsOpeCod, "D")
If rs.EOF Then
   MsgBox "No se asignó Cuenta de Regularización a Operación", vbInformation, "¡Aviso!"
   RSClose rs
   Exit Sub
End If
sCtaPend = rs!cCtaContCod

txtCtaCod.rs = oOpe.CargaOpeCtaArbol(gsOpeCod, "H")
CargaLista gsCodAge
txtFecha = gdFecSis
Set oOpe = Nothing
If Mid(gsOpeCod, 3, 1) = "2" Then
    lblSimbolo.ForeColor = gsColorME
    txtImporte.ForeColor = gsColorME
    txtImporte.BackColor = gsBackColorME
    gsSimbolo = gcME
Else
    lblSimbolo.ForeColor = gsColorMN
    txtImporte.ForeColor = gsColorMN
    txtImporte.BackColor = gsBackColorMN
    gsSimbolo = gcMN
End If
lblSimbolo = gsSimbolo
Dim oAreas As DActualizaDatosArea
Set oAreas = New DActualizaDatosArea
txtAgeCod.rs = oAreas.GetAgencias
Set oAreas = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not glAceptar Then
   If MsgBox(" ¿ Seguro que desea salir sin Grabar ?", vbQuestion + vbYesNo, "Confirmación!") = vbNo Then
      Cancel = 1
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oContFunct = Nothing
End Sub

Private Sub lstPago_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstPago.SortKey = ColumnHeader.Index - 1
lstPago.Sorted = True
End Sub

Private Sub lstPago_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtMovDesc = Item.SubItems(5)
txtImporte = Format(Item.SubItems(2), gsFormatoNumeroView)
End Sub

Private Sub lstPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtAgeCod = "" Then
      txtAgeCod.SetFocus
   Else
      txtCtaCod.SetFocus
   End If
End If
End Sub

Private Sub txtAgeCod_EmiteDatos()
    lblAgeDesc.Caption = txtAgeCod.psDescripcion
'    ProgressShow oPrg, Me
    
    CargaLista txtAgeCod
'    ProgressClose oPrg, Me
    If txtAgeCod.Visible Then
      If lblAgeDesc <> "" Then
         If lstPago.ListItems.Count > 0 Then
            lstPago_ItemClick lstPago.ListItems(lstPago.SelectedItem.Index)
            lstPago.SetFocus
         End If
      End If
    End If
End Sub

Private Sub txtCtaCod_EmiteDatos()
If txtCtaCod.psDescripcion <> "" Then
    If txtCtaCod.Enabled Then
        txtCtaDes = txtCtaCod.psDescripcion
        AsignaCtaObj txtCtaCod
        If txtCtaCod.Text <> "" Then
            txtMovDesc.SetFocus
        Else
            txtCtaDes = ""
        End If
    Else
        txtCtaCod.Text = ""
        txtCtaCod.Enabled = True
    End If
End If
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   lstPago.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   CmdAceptar.SetFocus
End If
End Sub

Public Sub AsignaCtaObj(ByVal psCtaContCod As String)
Dim lnItem As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
lnItem = 1
EliminaFgObj lnItem
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Case ObjDescomEfectivo
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case ObjPersona
                Set rs = Nothing
            Case Else
                lsRaiz = "Varios"
                Set rs = GetObjetos(Val(rs1!cObjetoCod))
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            AdicionaObj psCtaContCod, lnItem, rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                        Else
                            txtCtaCod.Text = ""
                            Exit Do
                        End If
                    Else
                        AdicionaObj psCtaContCod, lnItem, rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    AdicionaObj psCtaContCod, lnItem, rs1!nCtaObjOrden, _
                                    UP.sPersCod, UP.sPersNombre, _
                                    lsFiltro, rs1!cObjetoCod
                End If
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Sub
Private Sub AdicionaObj(sCodCta As String, nFila As Integer, _
                        psOrden As String, psObjetoCod As String, psObjDescripcion As String, _
                        psSubCta As String, psObjPadre As String)
Dim nItem As Integer
    fgObj.AdicionaFila
    nItem = fgObj.Row
    fgObj.TextMatrix(nItem, 0) = nFila
    fgObj.TextMatrix(nItem, 1) = psOrden
    fgObj.TextMatrix(nItem, 2) = psObjetoCod
    fgObj.TextMatrix(nItem, 3) = psObjDescripcion
    fgObj.TextMatrix(nItem, 4) = sCodCta
    fgObj.TextMatrix(nItem, 5) = psSubCta
    fgObj.TextMatrix(nItem, 6) = psObjPadre
    fgObj.TextMatrix(nItem, 7) = nFila
    'fgDetalle.TextMatrix(fgDetalle.Row, 6) = psObjetoCod
    
End Sub

Private Sub EliminaFgObj(nItem As Integer)
Dim K  As Integer, m As Integer
K = 1
Do While K < fgObj.Rows
   If Len(fgObj.TextMatrix(K, 1)) > 0 Then
      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
         fgObj.EliminaFila K, False
      Else
         K = K + 1
      End If
   Else
      K = K + 1
   End If
Loop
End Sub

