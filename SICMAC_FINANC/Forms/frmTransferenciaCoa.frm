VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransferenciaCoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Información al COA"
   ClientHeight    =   4740
   ClientLeft      =   3525
   ClientTop       =   1575
   ClientWidth     =   5265
   Icon            =   "frmTransferenciaCoa.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5265
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      ForeColor       =   &H8000000D&
      Height          =   2085
      Left            =   270
      ScaleHeight     =   2055
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   1950
      Width           =   465
      Begin VB.Image Image1 
         Height          =   390
         Left            =   30
         Picture         =   "frmTransferenciaCoa.frx":030A
         Stretch         =   -1  'True
         Top             =   15
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo"
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
      Height          =   765
      Left            =   270
      TabIndex        =   8
      Top             =   180
      Width           =   4785
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmTransferenciaCoa.frx":064C
         Left            =   2610
         List            =   "frmTransferenciaCoa.frx":0674
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1965
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   870
         TabIndex        =   0
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         Left            =   390
         TabIndex        =   10
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   "Mes"
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
         Left            =   2070
         TabIndex        =   9
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formato"
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
      Height          =   735
      Left            =   270
      TabIndex        =   7
      Top             =   1020
      Width           =   2265
      Begin VB.CheckBox chkTexto 
         Caption         =   "Texto"
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
         Height          =   285
         Left            =   1170
         TabIndex        =   3
         Top             =   315
         Width           =   870
      End
      Begin VB.CheckBox chkDBF 
         Caption         =   "DBF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   390
      Left            =   3660
      TabIndex        =   6
      Top             =   4200
      Width           =   1410
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "&Transferir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   1410
   End
   Begin VB.ListBox lstCoa 
      Height          =   2085
      ItemData        =   "frmTransferenciaCoa.frx":06DC
      Left            =   750
      List            =   "frmTransferenciaCoa.frx":06EC
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1950
      Width           =   4320
   End
End
Attribute VB_Name = "frmTransferenciaCoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Long
Dim SQLCoa As String
Dim rs    As ADODB.Recordset
Dim rsCoa As ADODB.Recordset
Dim lsAnio As String
Dim lsMes As String
Dim lsPeriodo As String
Dim dbCoa  As ADODB.Connection
Dim oCoa   As NContImpreReg
Dim oBarra As clsProgressBar
Dim sConCoa As String

Private Sub TransfiereNotasDBF()
Dim SqlNot As String
Dim lsFechaDoc  As String
Dim lsFechaDocREF As String
Dim ldFecha As Date
Dim sDocRef As String
Dim sTipRef As String
Dim scdoctpo As String
Dim sSerie As String
Dim sNumero As String
Dim nMovOtroImporte As Integer

On Error GoTo ERROR
lsAnio = txtAnio.Text
lsMes = Format(Me.cboMes.ListIndex + 1, "00")
lsPeriodo = lsAnio & lsMes
gsOpeCod = "700101"
'**Modificado por ALPA 26032008*******************************************************
'Set rs = oCoa.CargaCOADocumentos(lsPeriodo, gcCtaIGV, gsOpeCod, "'7','8'")
Set rs = oCoa.CargaCOADocumentos(lsPeriodo, gcCtaIGV, gsOpeCod, "7,8")
'**End********************************************************************************
If rs.EOF And rs.BOF Then
   MsgBox "No Existen Notas de Crédito ni Débito para Transferir", vbInformation, "¡Aviso!"
Else
   oBarra.ShowForm Me
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Max = rs.RecordCount
   dbCoa.BeginTrans
   I = 1
   Do While Not rs.EOF
      'ldFecha = "0"
      ldFecha = ""
      sDocRef = ""
      sTipRef = ""
      scdoctpo = ""
      sSerie = ""
      sNumero = ""
      nMovOtroImporte = 0
      'ldFecha = IIf(IsNull(rs!FecRef), 0, rs!FecRef)
      'ldFecha = IIf(IsNull(rs!cMovNro), 0, Format(Left(rs!cMovNro, 8), "yyyymmdd"))
      lsFechaDoc = Format(Year(rs!dDocFecha), "0000") & Format(Month(rs!dDocFecha), "00") & Format(Day(rs!dDocFecha), "00")
      lsFechaDocREF = Format(Year(ldFecha), "0000") & Format(Month(ldFecha), "00") & Format(Day(ldFecha), "00")
         '**Modificado por ALPA 26032008****************************************
         '**************************************************************************
         SQLCoa = "INSERT INTO X68CPCOA(X68NROREG,X68IDENT,X68PERIODO,X68FECHA,X68TIPO,X68SERIE," _
         & " X68NUMERO,X68BASE,X68IGV,X68CODIGO,X68FLAGMON,X68NORDEN,X68TIPORI,X68SERORI,X68NUMORI," _
         & " X68FECORI,X68ESTADO) VALUES( " & I & " ,'" & IIf(IsNull(rs!cProvRuc), " ", rs!cProvRuc) & "','" & lsPeriodo & "'," _
         & " '" & lsFechaDoc & "' , '" & scdoctpo & "','" & sSerie & "','" & sNumero & "'," _
         & " " & IIf(Mid(rs!cOpeCod, 3, 1) = 2, rs!nMovMEImporte - nMovOtroImporte, rs!nMovImporte - nMovOtroImporte) & "," _
         & " " & nMovOtroImporte & ", '003','" & IIf(Mid(rs!cOpeCod, 3, 1) = 2, 1, 0) & " '," _
         & " '" & Format(I, "000000") & "','" & sTipRef & "','" & Mid(IIf(IsNull(sDocRef), "", sDocRef), 1, 3) & "' ," _
         & " '" & Mid(IIf(IsNull(sDocRef), "", sDocRef), 5, 15) & "','" & lsFechaDocREF & "','1')"
         
         'SQLCoa = "INSERT INTO X68CPCOA(X68NROREG,X68IDENT,X68PERIODO,X68FECHA,X68TIPO,X68SERIE," _
         '& " X68NUMERO,X68BASE,X68IGV,X68CODIGO,X68FLAGMON,X68NORDEN,X68TIPORI,X68SERORI,X68NUMORI," _
         '& " X68FECORI,X68ESTADO) VALUES( " & I & " ,'" & IIf(IsNull(rs!cProvRuc), " ", rs!cProvRuc) & "','" & lsPeriodo & "'," _
         '& " '" & lsFechaDoc & "' , '" & rs!cdoctpo & "','" & rs!Serie & "','" & rs!Numero & "'," _
         '& " " & IIf(Mid(rs!cOpeCod, 3, 1) = 2, rs!nMovMEImporte - rs!nMovOtroImporte, rs!nMovImporte - rs!nMovOtroImporte) & "," _
         '& " " & rs!nMovOtroImporte & ", '003','" & IIf(Mid(rs!cOpeCod, 3, 1) = 2, 1, 0) & " '," _
         '& " '" & Format(I, "000000") & "','" & rs!TipRef & "','" & Mid(IIf(IsNull(rs!DocRef), "", rs!DocRef), 1, 3) & "' ," _
         '& " '" & Mid(IIf(IsNull(sDocRef), "", rs!DocRef), 5, 15) & "','" & lsFechaDocREF & "','1')"
         '**************************************************************************
         '**End*********************************************************************
         
      dbCoa.Execute SQLCoa
      I = I + 1
      oBarra.Progress I, "TRANSFERENCIA DE N/C y N/D", "", "Transfiriendo...", vbBlue
      rs.MoveNext
   Loop
   dbCoa.CommitTrans
   oBarra.CloseForm Me
End If
RSClose rs

'Set RS = Nothing
Exit Sub
ERROR:
    MsgBox TextErr(Err.Description)
    dbCoa.RollbackTrans
End Sub
Private Sub TransfiereProveeDBF()
Dim lsNOmbre As String
lsAnio = txtAnio.Text
lsMes = Format(Me.cboMes.ListIndex + 1, "00")
lsPeriodo = lsAnio & lsMes
Set rs = oCoa.CargaCOAProveedores(lsPeriodo)
Set rsCoa = New ADODB.Recordset
If RSVacio(rs) Then
   MsgBox "No existen Proveedores para Transferir", vbInformation, "¡Aviso!"
Else
   oBarra.ShowForm Me
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Max = rs.RecordCount
   I = 0
   Do While Not rs.EOF
      lsNOmbre = Replace(IIf(IsNull(rs!cNomPers), "", rs!cNomPers), "'", " ", , , vbTextCompare)
      SQLCoa = "SELECT Count(X61NRORUC) Valor FROM X61COAPR WHERE X61NRORUC = '" & Trim(rs!cProvRuc) & "'"
      If rsCoa.State = adStateOpen Then rsCoa.Close
      rsCoa.Open SQLCoa, dbCoa
      'If rsCoa.EOF Then
      If rsCoa!Valor = 0 Then
         SQLCoa = "INSERT INTO X61COAPR(X61NROREG,X61NRORUC,X61CINTER,X61NOMBRE,X61ESTADO)" _
         & " VALUES(" & I + 1 & ",'" & rs!cProvRuc & "','','" & lsNOmbre & "','0')"
         dbCoa.Execute SQLCoa
      End If
      I = I + 1
      oBarra.Progress I, "TRANSFERENCIA DE PROVEEDORES", "", "Transfiriendo...", vbBlue
      rs.MoveNext
   Loop
   If rsCoa.State = adStateOpen Then rsCoa.Close: Set rsCoa = Nothing
   oBarra.CloseForm Me
End If
RSClose rs

Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Private Sub TransfiereComprobanteDBF()
Dim SQLCom As String
Dim lnImporte As String
Dim lsFechaDoc As String
Dim lsImpre As String
Dim lnLin   As Integer
Dim P       As Integer
Dim nIGV As Currency
Dim nOtros As Currency
Dim lsTipoGravado As String
Dim lnTotal As Currency
Dim lnTotIGV As Currency
Dim lnTotOtros As Currency
Dim lnMovTot As Currency
On Error GoTo ERROR
lsAnio = txtAnio.Text
lsMes = Format(Me.cboMes.ListIndex + 1, "00")
lsPeriodo = lsAnio & lsMes
gsOpeCod = "700101"


Set rs = oCoa.CargaCOADocumentos(lsPeriodo, gcCtaIGV, gsOpeCod)
If rs.EOF Then
    MsgBox "No existen Comprobantes para Migrar", vbInformation, "¡Aviso!"
    Exit Sub
Else
   oBarra.ShowForm Me
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Max = rs.RecordCount
   I = 0
   P = 0
   lnLin = gnLinPage
   Do While Not rs.EOF
    If rs!nDocTpo = 14 And rs!nIGV = 0 Then
    Else
       I = I + 1
      nOtros = 0
      nIGV = 0
      oBarra.Progress rs.Bookmark, "TRANSFERENCIA DE COMPROBANTES", "", "Generando Reporte...", vbBlue
      If lnLin > gnLinPage - 6 Then
          lnLin = 0
          Linea lsImpre, Cabecera("TRANSFERENCIA DE COMPROBANTES", P, gsSimbolo, gnColPage, "", Format(gdFecSis, gsFormatoFechaView)), , lnLin
          lnLin = lnLin + 5
          Linea lsImpre, String(75, "="), , lnLin
          Linea lsImpre, " DOCUMENTO                          PROVEEDOR                        RUC                  V.Venta      I.G.V.       Otros        P.Venta ", , lnLin
          Linea lsImpre, String(75, "-"), , lnLin
      End If
      If Mid(rs!cOpeCod, 3, 1) = "2" Then
         If rs!nOtros = 0 And rs!nIGV <> 0 Then
            nIGV = Round(rs!nMovImporte - (rs!nMovImporte / (1 + gnIGVValor)), 2)
         Else
            If rs!nIGV <> 0 Then
               nIGV = Round(rs!nIGV * rs!nMovImporte / rs!nMovMEImporte, 2)
               nOtros = Round(rs!nOtros * rs!nMovImporte / rs!nMovMEImporte, 2)
            End If
        End If
      Else
         nIGV = rs!nIGV
         nOtros = Round(rs!nOtros * rs!nTC, 2)
      End If
      'lnTotal = lnTotal + (rs!nMovImporte - nIGV - nOtros) original gitu
      lnTotal = lnTotal + (rs!nMovImporte - nIGV - nOtros)
      lnTotIGV = lnTotIGV + nIGV
      lnTotOtros = lnTotOtros + nOtros
      lnMovTot = lnMovTot + rs!nMovImporte
      Linea lsImpre, Justifica(Str(I), 5) & Justifica(rs!nDocTpo, 3) & Justifica(rs!cDocNro, 15) & Justifica(rs!dDocFecha, 11) & Justifica(rs!cPersNombre, 40) & Justifica(rs!cProvRuc, 11) & PrnVal(rs!nMovImporte - nIGV - nOtros, 13, 2) & " " & PrnVal(nIGV, 12, 2) & " " & PrnVal(nOtros, 12, 2) & PrnVal(rs!nMovImporte, 14, 2), , lnLin
    End If
    rs.MoveNext
   Loop
   Linea lsImpre, Space(85) & PrnVal(lnTotal, 13, 2) & " " & PrnVal(lnTotIGV, 12, 2) & " " & PrnVal(lnTotOtros, 12, 2) & PrnVal(lnMovTot, 14, 2), , lnLin
   oBarra.CloseForm Me
   EnviaPrevio lsImpre, "TRANSFERENCIA DE COMPROBANTES", gnLinPage, True
End If
If MsgBox(" ¿ Desea Transferir Documentos ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
    RSClose rs
    Exit Sub
End If
rs.MoveFirst
If Not rs.EOF Then
   oBarra.ShowForm Me
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Max = rs.RecordCount
   dbCoa.BeginTrans
   I = 1
   Dim nPos As Integer
   Dim sSerie As String
   Dim sTexto As String
   sTexto = "  RUC  PROVEEDOR                      PERIODO     FECHA   TPO  NUMERO       V.VENTA     IGV " & oImpresora.gPrnSaltoLinea
   dbCoa.Execute "DELETE FROM x58CPCOA"
   'dbCoa.Execute "PACK"
   
   Do While Not rs.EOF
    nOtros = 0
    nIGV = 0

      If rs!nDocTpo = 14 And rs!nIGV = 0 Then
      Else
          nPos = InStr(rs!cDocNro, "-")
          sSerie = ""
          If nPos > 0 Then
             sSerie = Mid(rs!cDocNro, 1, nPos - 1)
          End If
          If Mid(rs!cOpeCod, 3, 1) = "2" Then
               If rs!nOtros = 0 And rs!nIGV <> 0 Then
                  nIGV = Round(rs!nMovImporte - (rs!nMovImporte / (1 + gnIGVValor)), 2)
               Else
                  If rs!nIGV <> 0 Then
                     nIGV = Round(rs!nIGV * rs!nMovImporte / rs!nMovMEImporte, 2)
                     nOtros = Round(rs!nOtros * rs!nMovImporte / rs!nMovMEImporte, 2)
                  End If
              End If
          Else
               nIGV = rs!nIGV
               nOtros = Round(rs!nOtros * rs!nTC, 2)
          End If
          lsTipoGravado = "003"
          If nIGV = 0 Then
            lsTipoGravado = "004"
          End If
          lsFechaDoc = Format(Year(rs!dDocFecha), "0000") & Format(Month(rs!dDocFecha), "00") & Format(Day(rs!dDocFecha), "00")
          SQLCoa = " INSERT INTO X58CPCOA (X58NROREG,X58IDENT,X58PERIODO,X58FECHA,X58TIPO,X58SERIE," _
          & " X58NUMERO,X58BASE,X58IGV,X58CODIGO,X58FLAGMON,X58NORDEN,X58ESTADO)" _
          & " VALUES(" & I & ",'" & rs!cProvRuc & "' ,'" & lsPeriodo & "','" & lsFechaDoc & "'," _
          & " '" & Format(rs!nDocTpo, "00") & "','" & sSerie & "','" & Mid(rs!cDocNro, nPos + 1, 20) & "'," & rs!nMovImporte - nIGV & " ," _
          & "  " & nIGV & ", '" & lsTipoGravado & "'," _
          & " '" & IIf(Mid(rs!cOpeCod, 3, 1) = 2, 0, 1) & " ','" & Format(I, "000000") & "','0')"
          dbCoa.Execute SQLCoa
          If nOtros <> 0 Then
             lsTipoGravado = "004"
             SQLCoa = " INSERT INTO X58CPCOA (X58NROREG,X58IDENT,X58PERIODO,X58FECHA,X58TIPO,X58SERIE," _
             & " X58NUMERO,X58BASE,X58IGV,X58CODIGO,X58FLAGMON,X58NORDEN,X58ESTADO)" _
             & " VALUES(" & I & ",'" & rs!cProvRuc & "' ,'" & lsPeriodo & "','" & lsFechaDoc & "'," _
             & " '" & Format(rs!nDocTpo, "00") & "','" & sSerie & "','" & Mid(rs!cDocNro, nPos + 1, 20) & "', " & nOtros & ", " _
             & " 0, '" & lsTipoGravado & "'," _
             & " '" & IIf(Mid(rs!cOpeCod, 3, 1) = 2, 0, 1) & " ','" & Format(I, "000000") & "','0')"
             dbCoa.Execute SQLCoa
          End If
          I = I + 1
      End If
      oBarra.Progress rs.Bookmark, "TRANSFERENCIA DE COMPROBANTES", "", "Transfiriendo...", vbBlue
      rs.MoveNext
   Loop
   oBarra.CloseForm Me
   dbCoa.CommitTrans
   MsgBox "Transferencia Completada con Exito", vbInformation, "Información"
Else
   MsgBox "No existen Comprobantes para Transferir", vbInformation, "¡Aviso!"
End If
RSClose rs
Exit Sub
ERROR:
    MsgBox TextErr(Err.Description)
'    dbCoa.RollbackTrans
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   lstCoa.SetFocus
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTransferir_Click()
Dim lbFlag As Boolean
Set oBarra = New clsProgressBar
lbFlag = False
If lstCoa.Selected(0) = True Then
    TransfiereProveeDBF
    lbFlag = True
End If
If lstCoa.Selected(1) = True Then
    TransfiereComprobanteDBF
    lbFlag = True
End If
If lstCoa.Selected(2) = True Then
    TransfiereNotasDBF
    lbFlag = True
End If
If lstCoa.Selected(3) = True Then
   MsgBox "No existen Exportaciones para Transferir", vbInformation, "¡Aviso!"
   lbFlag = True
End If
Set oBarra = Nothing
If lbFlag = False Then MsgBox "Selecione un Opción para ser Trasferida"
End Sub


Private Sub Form_Load()
On Error GoTo ConexionErr
 sConCoa = "DSN=DSNCoa;"
 Set dbCoa = New ADODB.Connection
 Set oCoa = New NContImpreReg
 CentraForm Me
    dbCoa.CommandTimeout = 30
    dbCoa.ConnectionTimeout = 30
    dbCoa.Open sConCoa
    txtAnio.Text = Str(Year(gdFecSis))
    cboMes.ListIndex = Month(gdFecSis) - 1
    RotateText 90, Picture1, "Times New Roman", 15, 25, 1500, "C O A"
   Exit Sub
ConexionErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
dbCoa.Close
Set dbCoa = Nothing
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboMes.SetFocus
End If
End Sub
