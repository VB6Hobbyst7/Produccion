VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmación de Recepción de Cheques"
   ClientHeight    =   6675
   ClientLeft      =   1350
   ClientTop       =   1590
   ClientWidth     =   8385
   Icon            =   "frmTransCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   2550
      Left            =   75
      TabIndex        =   14
      Top             =   45
      Width           =   8190
      Begin VB.Frame Frame4 
         Height          =   960
         Left            =   5580
         TabIndex        =   18
         Top             =   1230
         Width           =   2475
         Begin VB.CheckBox chkCredito 
            Caption         =   "Sin Relacion Ahorros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   135
            TabIndex        =   5
            Top             =   150
            Width           =   2145
         End
         Begin VB.CommandButton cmdBuscar 
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
            Left            =   270
            TabIndex        =   6
            Top             =   510
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Height          =   990
         Left            =   5580
         TabIndex        =   16
         Top             =   150
         Width           =   2490
         Begin VB.CheckBox chkIngCaja 
            Caption         =   "Ingreso a Caja General"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   75
            TabIndex        =   3
            Top             =   165
            Width           =   2340
         End
         Begin MSMask.MaskEdBox txtIngCaja 
            Height          =   345
            Left            =   870
            TabIndex        =   4
            Top             =   465
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
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
            Caption         =   "Fecha :"
            Height          =   255
            Left            =   225
            TabIndex        =   17
            Top             =   510
            Width           =   645
         End
      End
      Begin VB.ListBox lstAgencias 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   810
         Width           =   3840
      End
      Begin VB.CommandButton cmdAgeTodas 
         Caption         =   "T&odas"
         Height          =   255
         Left            =   255
         TabIndex        =   0
         Top             =   495
         Width           =   1650
      End
      Begin VB.CommandButton cmdAgeNinguna 
         Caption         =   "Ninguna"
         Height          =   255
         Left            =   2175
         TabIndex        =   1
         Top             =   495
         Width           =   1650
      End
      Begin MSComctlLib.ProgressBar PBarra 
         Height          =   210
         Left            =   4005
         TabIndex        =   19
         Top             =   2250
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label LblAgen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   135
         TabIndex        =   20
         Top             =   2265
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE LAS AGENCIAS"
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
         Height          =   195
         Left            =   765
         TabIndex        =   15
         Top             =   210
         Width           =   2565
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3480
      Left            =   75
      TabIndex        =   13
      Top             =   2670
      Width           =   8205
      Begin VB.CheckBox chkCheques 
         Caption         =   "Seleccionar &Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   7
         Top             =   -15
         Width           =   1770
      End
      Begin MSComctlLib.ListView lvCheque 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Banco"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No. Cheque"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Persona"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha Valor."
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Estado"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Fecha Registro"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PerCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cMovNro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cDocTpo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Cuenta de Banco"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Cuenta de Cliente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "cCodBco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "cAgeCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Agencia"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "cCodBanco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "MontoChq"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Tipo Cambio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Moneda"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """S/.""#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   360
      Left            =   6255
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6210
      Width           =   1845
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   6240
      Width           =   1635
   End
   Begin VB.CommandButton cmdTransCheque 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Monto "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5505
      TabIndex        =   11
      Top             =   6270
      Width           =   630
   End
End
Attribute VB_Name = "frmTransCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim lSalir As Boolean
Dim lnMonto As Currency
Dim nTotal As Currency
Dim lsClave As String
Dim lnNumCheques As Long

Private Sub MuestraCheques(pAgencia As String, pNomAgencia As String)
Dim lvItem As ListItem
Dim sCta As String
Dim lrs As New ADODB.Recordset
On Error GoTo MuestraChequesErr
Dim oCon As DConecta
Set oCon = New DConecta
If gbBitCentral Then
       
    If chkCredito.value = 0 Then
  
        sSql = "SELECT  IsNull((Select cPersNombre From Persona PE where PE.cPersCod = DR.cPersCod),'') Banco, DRC.cPersCod, DRC.cIFTpo, "
        sSql = sSql & "      DRC.cCtaCod cCodCta, DRC.cNroDoc cNumChq, dbo.FechaMov(M.cMovNro) dRegChq, DR.dValorizacion dValorChq, "
        sSql = sSql & "      DR.nMonto nMontoChq, Right(M.cMovNro,4) cCodUsu, M.cOpeCod cCodOpe, CO.cConsDescripcion Estado, DR.cIFCta cCtaBco, "
        sSql = sSql & "      SubString(M.cMovNro,18,2) cCodAge, '' cNomPers, bPlaza cPlaza, cDepIF cDepBco, nEstado cEstChq, '' cValor1, "
        sSql = sSql & "      IsNull((Select cAgeDescripcion From Agencias AGE Where AGE.cAgeCod = SubString(M.cMovNro,18,2)),'') Agencia "
        sSql = sSql & " FROM DOcRecCapta DRC "
        sSql = sSql & "     Inner Join DocRec DR On DRC.nTpoDoc = DR.nTpoDoc And DRC.cNroDoc = DR.cNroDoc And DRC.cPersCod = DR.cPersCod And DRC.cIFTpo = DR.cIFTpo "
        sSql = sSql & "     Inner Join Mov M On M.nMovNro = DRC.nMovNro And M.cOpeCod in ('900031','200102','200105','200202','210102','210105','220102','220105','220202') "
        sSql = sSql & "     Inner Join Constante CO On CO.nConsCod = 1009 And CO.nConsvalor = DR.bPlaza "
        sSql = sSql & " WHERE Substring(DRC.cCtaCod,4,2) = '" & Right(pAgencia, 2) & "' "
        sSql = sSql & "     and DR.nEstado not in (3,4,5) And M.nMovflag = 0 And DRC.nTpoDoc = " & TpoDocCheque & " and nConfCaja IN (" & ChqCGSinConfirmacion & "," & ChqCGNoConfirmado & ") "
        sSql = sSql & "     AND (SELECT COUNT(DRR.cNroDoc) FROM DocRecRel DRR WHERE DRR.nTpoDoc=DR.nTpoDoc AND DRR.cNroDoc=DR.cNroDoc AND DRR.cPersCod=DR.cPersCod AND DRR.cIFTpo=DR.cIFTpo AND DRR.cIFCta=DR.cIFCta)=0 " 'EJVG20140415

    Else
        
        sSql = "SELECT DISTINCT IsNull((Select cPersNombre From Persona PE where PE.cPersCod = DR.cPersCod),'') Banco, "
        sSql = sSql & " DR.cPersCod, DR.cIFTpo, '' cCodCta, DR.cNroDoc cNumChq, dbo.FechaMov(M.cMovNro) dRegChq, "
        sSql = sSql & " DR.dValorizacion dValorChq, DR.nMonto nMontoChq, Right(M.cMovNro,4) cCodUsu, M.cOpeCod cCodOpe, "
        sSql = sSql & " CO.cConsDescripcion Estado, DR.cIFCta cCtaBco, "
        sSql = sSql & " SubString(M.cMovNro,18,2) cCodAge, '' cNomPers, bPlaza cPlaza, cDepIF cDepBco, nEstado cEstChq, '' cValor1, "
        sSql = sSql & " IsNull((    Select cAgeDescripcion "
        sSql = sSql & "             From Agencias AGE "
        sSql = sSql & "             Where Age.cAgeCod = substring(m.cMovNro, 18, 2) "
        sSql = sSql & "        ),'') Agencia "
        sSql = sSql & " From DocRec DR "
        sSql = sSql & " Inner Join MovDoc MD "
        sSql = sSql & "            On MD.cDocNro=DR.cNroDoc "
        sSql = sSql & " Inner Join Mov M on MD.nMovNro=M.nMovNro "
        sSql = sSql & " And M.cOpeCod in ('900031','200102','200105','200202','210102','210105','220102','220105','220202')"
        sSql = sSql & " Inner Join Constante CO On CO.nConsCod = 1009 And CO.nConsvalor = DR.bPlaza "
        sSql = sSql & " WHERE substring(M.cMovNro, 18,2) = '" & Right(pAgencia, 2) & "' "
        sSql = sSql & " and DR.nEstado not in (3,4,5) "
        sSql = sSql & " And M.nMovflag = 0 "
        sSql = sSql & " And MD.nDocTpo = " & TpoDocCheque & " "
        sSql = sSql & " and nConfCaja IN (" & ChqCGSinConfirmacion & "," & ChqCGNoConfirmado & ") "
        sSql = sSql & " and DR.cNroDoc + DR.cPersCod  + DR.cIFtpo NOT IN( "
        sSql = sSql & "     Select DRC.cNroDoc + DRC.cPersCod  + DRC.cIFtpo "
        sSql = sSql & "     FROM DocRecCapta  DRC "
        sSql = sSql & "     Where DRC.nTpoDoc = " & TpoDocCheque & " "
        sSql = sSql & "     ) "
        sSql = sSql & "AND (SELECT COUNT(DRR.cNroDoc) FROM DocRecRel DRR WHERE DRR.nTpoDoc=DR.nTpoDoc AND DRR.cNroDoc=DR.cNroDoc AND DRR.cPersCod=DR.cPersCod AND DRR.cIFTpo=DR.cIFTpo AND DRR.cIFCta=DR.cIFCta)=0 " 'EJVG20140415
        
    End If
    
    oCon.AbreConexion

Else
    sSql = " SELECT a.cCodCta,a.cEstChq, d.cBcoDes as Banco, a.cPlaza, a.dRegChq, d.cValor1 , a.cCodBco," _
         & "        a.cNumChq, a.dValorChq, " _
         & "        e.cNomTab as Estado, a.nMontoChq, a.cCtaBco, a.cDepBco, ISNULL(p.cNomPers ,'') cNomPers " _
         & " FROM Cheque a JOIN TransAho ta ON ta.dFecTran = a.dRegChq and ta.cCodCta = a.cCodCta and ta.cNumdoc = a.cNumChq " _
         & "          LEFT JOIN (SELECT cCodCta, MIN(cCodPers) cCodPers FROM PersCuenta WHERE cRelaCta = 'TI' GROUP BY cCodCta ) pc ON pc.cCodCta = ta.ccodCta  " _
         & "          LEFT JOIN DBPersona..Persona P ON P.cCodPers = pc.cCodPers," _
         & "      DBComunes..Bancos d, DBComunes..TablaCod e " _
         & " WHERE (cTransChq = '0' or cTransChq IS NULL) and (ta.cCodAge = '" & pAgencia & "' or a.cEstChq = 'T') and ta.cFlag IS NULL " _
         & "       and convert(int,a.cCodBco) = d.nBcoCod and a.cEstChq = e.cValor and nMonTran > 0 and " _
         & "       (SUBSTRING(e.cCodTab,1,2) = '13') AND a.cEstChq NOT IN('X') " _
         & "       and a.cEstChq NOT IN('A','R') " _
         & " ORDER BY d.cBcoDes "
    If oCon.AbreConexion Then ''Remota(Right(pAgencia, 2), True)
    Else
       Exit Sub
    End If
End If
   RSClose lrs
   Set lrs = oCon.CargaRecordSet(sSql)
   If Not lrs.EOF Then
        PBarra.Visible = True
        PBarra.Min = 0
        PBarra.Max = lrs.RecordCount
        PBarra.value = 0
        Screen.MousePointer = 11
        Do While Not lrs.EOF
            LblAgen.Caption = pNomAgencia
            DoEvents
            Set lvItem = lvCheque.ListItems.Add(, , Trim(lrs!banco))
            lvItem.SubItems(1) = lrs!cNumChq
            lvItem.SubItems(2) = lrs!cNomPers
            lvItem.SubItems(3) = Format(lrs!dValorChq, "dd/mm/yyyy hh:mm:ss")
            lvItem.SubItems(4) = Format(lrs!nMontoChq, gsFormatoNumeroView) 'Format(lrs!nMontoTC, gsFormatoNumeroView)
            lvItem.SubItems(5) = Trim(lrs!Estado)
            lvItem.SubItems(6) = Format(lrs!dRegChq, "dd/mm/yyyy hh:mm:ss")
            lvItem.SubItems(7) = lrs!cPlaza
            lvItem.SubItems(8) = lrs!cDepBco
            lvItem.SubItems(9) = lrs!cEstChq
            lvItem.SubItems(10) = Trim(lrs!cCtaBco)
            lvItem.SubItems(11) = Trim(lrs!cCodCta)
            If Not gbBitCentral Then
                If Not IsNull(lrs!cValor1) Then
                    lvItem.SubItems(12) = Right("00" & Trim(Str(lrs!cValor1)), 2)
                End If
                lvItem.SubItems(15) = Trim(lrs!cCodBco)
            Else
                lvItem.SubItems(15) = Trim(lrs!ciftpo) & "." & Trim(lrs!cPersCod)
            End If
            lvItem.SubItems(13) = pAgencia
            lvItem.SubItems(14) = pNomAgencia
            lvItem.SubItems(16) = Format(lrs!nMontoChq, gsFormatoNumeroView)
            lvItem.SubItems(17) = gnTipCambio 'Format(lrs!nTipCambio)
            lvItem.SubItems(18) = Mid(lrs!cCodCta, 6, 1)           'Format(lrs!CMONEDA)
            nTotal = nTotal + lrs!nMontoChq
            lnNumCheques = lnNumCheques + 1
            PBarra.value = PBarra.value + 1
            lrs.MoveNext
        Loop
        Screen.MousePointer = 0
        txtMonto = Format(nTotal, "#,#00.00")
        LblAgen.Caption = Str(lnNumCheques) & " Cheques Encontrados"
        PBarra.Visible = False
    End If
    RSClose lrs
    oCon.CierraConexion
Exit Sub
MuestraChequesErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    
End Sub

Private Sub chkCheques_Click()
If Me.chkCheques.value = 1 Then
    Marcar True
Else
    Marcar False
End If
End Sub

Private Sub chkIngCaja_Click()
If Me.chkIngCaja.value = 1 Then
    Me.txtIngCaja.Enabled = True
Else
    Me.txtIngCaja.Enabled = False
End If
End Sub

Private Sub cmdAgeNinguna_Click()
MarcarLista False
End Sub

Private Sub cmdAgeTodas_Click()
MarcarLista True
End Sub

Private Sub cmdbuscar_Click()
Dim rs As New ADODB.Recordset
Dim lsAgencia As String
Dim CadCon As String
Dim I As Integer

cmdBuscar.Enabled = False
nTotal = 0
Me.txtMonto = "0.00"
lvCheque.ListItems.Clear
lnNumCheques = 0
LblAgen.Caption = ""
For I = 0 To Me.lstAgencias.ListCount - 1
   If lstAgencias.Selected(I) = True Then
        lsAgencia = gsCodCMAC & Right(lstAgencias.List(I), 2)
        MuestraCheques lsAgencia, Mid(lstAgencias.List(I), 1, 25)
   End If
Next
cmdBuscar.Enabled = True
LblAgen.Caption = ""
If Me.lvCheque.ListItems.Count = 0 Then MsgBox "No se encontrarón Cheques en las Agencias Selecionadas", vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub MarcarLista(Valor As Boolean)
Dim I As Integer
 For I = 0 To Me.lstAgencias.ListCount - 1
     lstAgencias.Selected(I) = Valor
 Next
End Sub

Sub Marcar(Valor As Boolean)
Dim I As Integer
 For I = 1 To Me.lvCheque.ListItems.Count
     lvCheque.ListItems(I).Checked = Valor
     If Valor = False Then
        Me.txtMonto.Text = "0.00"
        lnMonto = 0
     Else
      Me.txtMonto.Text = Format(nTotal, "#,#00.00")
        lnMonto = nTotal
     End If
 Next
End Sub

Private Sub cmdTransCheque_Click()
Dim sql As String
Dim I As Integer
Dim lsNumChq As String
Dim lsFecVal As String
Dim lnMonto As String
Dim lsEstado As String
Dim lsFecReg As String
Dim lsPlaza As String
Dim lsDepBco As String
Dim lsCtaBco As String
Dim lsCuenta As String
Dim lsCodBco As String
Dim lsCodUsu As String
Dim lsCodAge As String
Dim lnTipCambio As String
Dim lnMontoTC As String
Dim lsMoneda As String

Dim lbFlag As Boolean
Dim lsMovNro    As String
Dim lsMovNroReg As String
Dim lnMovNroReg As Long
Dim nCont    As Integer
Dim CadCon As String
Dim ContList As Integer
Dim lbTrans As Boolean
Dim oCon As New DConecta

On Error GoTo ERROR
lbTrans = False
For I = 1 To lvCheque.ListItems.Count
 If lvCheque.ListItems(I).Checked = True Then
    lbFlag = True
    Exit For
  Else
    lbFlag = False
 End If
Next
If Me.chkIngCaja.value = 1 Then
    If ValFecha(Me.txtIngCaja) = False Then
        Exit Sub
    End If
End If
If lbFlag = False Then
  MsgBox "No selecciono nigun Cheque para ser transferido", vbInformation, "Aviso"
  Exit Sub
End If

If MsgBox("Desea Confirmar recepción de Cheques Seleccionados?", vbYesNo + vbDefaultButton2 + vbQuestion, "Aviso") = vbNo Then
   Exit Sub
End If
Dim oMov As DMov
Set oMov = New DMov
lbFlag = False
nCont = 0
oMov.BeginTrans

lbTrans = True
lsMovNro = ""
For I = 1 To Me.lvCheque.ListItems.Count
   If lvCheque.ListItems(I).Checked = True Then
        nCont = nCont + 1
        lsNumChq = Trim(lvCheque.ListItems(I).SubItems(1))
        lsFecVal = lvCheque.ListItems(I).SubItems(3)
        lnMonto = Format(lvCheque.ListItems(I).SubItems(16), "#0.00")
        lsEstado = lvCheque.ListItems(I).SubItems(9)
        If chkIngCaja.value = 0 Then
            lsFecReg = lvCheque.ListItems(I).SubItems(6)
        Else
            lsFecReg = Format(txtIngCaja, "dd/mm/yyyy")
        End If
        lsPlaza = "1"
        If lvCheque.ListItems(I).SubItems(7) = "1" Then
            lsPlaza = "0"
        End If
        
        lsDepBco = lvCheque.ListItems(I).SubItems(8)
        lsCtaBco = Trim(lvCheque.ListItems(I).SubItems(10))
        lsCuenta = lvCheque.ListItems(I).SubItems(11)
        lsCodBco = Format(lvCheque.ListItems(I).SubItems(12), "00")
        lsCodUsu = gsCodUser
        lsCodAge = lvCheque.ListItems(I).SubItems(13)
                 
        lnTipCambio = lvCheque.ListItems(I).SubItems(17)
        lnMontoTC = Format(lvCheque.ListItems(I).SubItems(4), "#0.00")
        lsMoneda = lvCheque.ListItems(I).SubItems(18)
        
        Dim psPersCodIF As String
        Dim psOpeCod    As String
        Dim psOpeRegChq As String
        
        Dim lsTpoIf As String
        Dim oIF As DCajaCtasIF
        Set oIF = New DCajaCtasIF
        If gbBitCentral Then
            psPersCodIF = Mid(lvCheque.ListItems(I).SubItems(15), 4, 13)
            lsTpoIf = Left(lvCheque.ListItems(I).SubItems(15), 2)
        Else
            lsTpoIf = "0101"
            If InStr(lvCheque.ListItems(I).Text, "RURAL") > 0 Then
                lsTpoIf = "0104"
            End If
            psPersCodIF = oIF.GetCodPersAuxIF(lsTpoIf & lsCodBco)
        End If
        If psPersCodIF = "" Then
            MsgBox lvCheque.ListItems(I).Text & " no existe como Institución Financiera", vbInformation, "¡Aviso!"
        Else
            If gbBitCentral Then
                If Mid(lsCuenta, 9, 1) = "2" Then
                    psOpeCod = gOpeCGOpeBancosRecibeChqAgMe
                    psOpeRegChq = gOpeCGOpeBancosRegChequesME
                Else
                    psOpeCod = gOpeCGOpeBancosRecibeChqAgMn
                    psOpeRegChq = gOpeCGOpeBancosRegChequesMN
                End If
            Else
                If Mid(lsCuenta, 6, 1) = "2" Then
                    psOpeCod = gOpeCGOpeBancosRecibeChqAgMe
                    psOpeRegChq = gOpeCGOpeBancosRegChequesME
                Else
                    psOpeCod = gOpeCGOpeBancosRecibeChqAgMn
                    psOpeRegChq = gOpeCGOpeBancosRegChequesMN
                End If
            End If
            glAceptar = True
            sql = "SELECT dr.*, md.dDocFecha FROM DocRec dr JOIN MovDoc md ON md.cDocNro = dr.cNroDoc and md.nDocTpo = dr.nTpoDoc WHERE dr.nTpoDoc = " & TpoDocCheque & " and dr.cNroDoc = '" & lsNumChq & "' and cPersCod = '" & psPersCodIF & "' and cIFCta = '" & lsCtaBco & "'  "
            Set rs = oMov.CargaRecordSet(sql)
            If rs.EOF And Not gbBitCentral Then   'Cheque aun no registrado
                oMov.InsertaCheque TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), lsPlaza, lsCtaBco, lnMonto, CDate(lsFecVal), _
                           CDate(lsFecVal), Val(lsEstado), ChqCGConfirmado, Val(lsMoneda), , "013", Trim(Right(lsCodAge, 2))
             
                'Movimiento de Confirmacion del Cheque
                lsMovNro = oMov.GeneraMovNro(txtIngCaja, Right(gsCodAge, 2), gsCodUser, lsMovNro, nCont)
                oMov.InsertaMov lsMovNro, psOpeCod, "Transferencia de Cheques de Agencia ", gMovEstContabNoContable, gMovFlagVigente
                gnMovNro = oMov.GetnMovNro(lsMovNro)
                oMov.InsertaMovDoc gnMovNro, TpoDocCheque, lsNumChq, Format(CDate(lsFecReg), gsFormatoFecha)
                oMov.InsertaChequeEstado TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), CDate(lsFecReg), gChqEstEnValorizacion, lsMovNro, lsCtaBco
                Dim lsObj As String
                If Mid(lsCuenta, 3, 2) = "23" Then
                    lsObj = "6001" & Mid(lsCuenta, 3, 3)
                Else
                    lsObj = "6002" & Mid(lsCuenta, 3, 3)
                End If
                oMov.InsertaMovObj gnMovNro, 1, 1, lsObj
            
            Else
                glAceptar = False
                If Not gbBitCentral Then
                    If Not CDate(Format(lsFecReg, gsFormatoFechaView)) = rs!dDocFecha Then
                        If MsgBox("Cheque Nro. " & lsNumChq & " fue registrado el " & rs!dValorizaRef & ". ¿ Desea Continuar ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbYes Then
                            glAceptar = True
                        End If
                    End If
                    If rs!cDepIF = gCGEstadosChqRechazado Or rs!nestado = gsChqEstExtornado Or glAceptar Then
                        oMov.ActualizaCheque TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), , , , lsCtaBco, lnMonto, gCGEstadosChqRecibido, , , gChqEstRegistrado
                        lsMovNro = oMov.GeneraMovNro(txtIngCaja, Right(gsCodAge, 2), gsCodUser, lsMovNro, nCont)
                        oMov.InsertaMov lsMovNro, psOpeCod, "Transferencia de Cheques de Agencia ", gMovEstContabNoContable, gMovFlagVigente
                        gnMovNro = oMov.GetnMovNro(lsMovNro)
                        oMov.InsertaMovDoc gnMovNro, TpoDocCheque, lsNumChq, Format(CDate(lsFecReg), gsFormatoFecha)
                        oMov.InsertaChequeEstado TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), CDate(lsFecReg), gChqEstEnValorizacion, lsMovNro, lsCtaBco
                        glAceptar = True
                    Else
                        oMov.ActualizaCheque TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), , , , lsCtaBco, rs!nMonto + lnMonto
                        glAceptar = True
                    End If
                Else
                    glAceptar = True
                End If
            End If
            If glAceptar Then
                If gbBitCentral Then
                    oCon.AbreConexion
                    'Actualizar DocRecCapta
                    sql = "Update DocRecCapta Set cIFCta = '" & Trim(lsCtaBco) & "' where cCtaCod ='" & Trim(lsCuenta) & "' and  cNroDoc='" & Trim(lsNumChq) & "' and cPersCod = '" & psPersCodIF & "' and nMonto = " & lnMonto & " "
                    oCon.Ejecutar sql
                    sql = "Update DocRecEst Set cIFCta = '" & Trim(lsCtaBco) & "' where cNroDoc='" & Trim(lsNumChq) & "' and cPersCod = '" & psPersCodIF & "' and cIFCta is NULL "
                    oCon.Ejecutar sql
                    oMov.ActualizaCheque TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), , , , lsCtaBco, lnMonto, gCGEstadosChqRecibido, ChqCGConfirmado
                Else
                    sql = "SELECT * FROM DocRecCapta WHERE nTpoDoc = " & TpoDocCheque & " and cNroDoc = '" & lsNumChq & "' and cPersCod = '" & psPersCodIF & "' and cCtaCod = '" & gsCodCMAC & lsCuenta & "' and nMonto = " & lnMonto & " and cIFCta = '" & Trim(lsCtaBco) & "' "
                    Set rs = oMov.CargaRecordSet(sql)
                    If rs.EOF Then
                        oMov.InsertaDocRecCapta TpoDocCheque, lsNumChq, psPersCodIF, Right(lsTpoIf, 2), gsCodCMAC & lsCuenta, lnMonto, lsCtaBco
                    End If
                    RSClose rs
                    If oCon.AbreConexion Then 'Remota(Right(lsCodAge, 2), True)
                        sql = "Update Cheque Set cTransChq = '1' where cCodCta ='" & Trim(lsCuenta) & "' and  cNumChq='" & Trim(lsNumChq) & "' and cCodBco='" & Trim(lvCheque.ListItems(I).SubItems(15)) & "' and cCtaBco='" & Trim(lsCtaBco) & "'"
                        oCon.Ejecutar sql
                    End If
                    oCon.CierraConexion
                End If
            End If
        End If
    End If
Next

ContList = lvCheque.ListItems.Count
I = 1
Do While I <= ContList
    If lvCheque.ListItems(I).Checked Then
        lvCheque.ListItems.Remove lvCheque.ListItems(I).Index
        ContList = ContList - 1
        I = I - 1
    End If
    I = I + 1
Loop
oMov.CommitTrans
Set oMov = Nothing
lbTrans = False

MsgBox "Transferecia Completada con Exito", vbOKOnly + vbInformation, "Aviso"
Exit Sub
ERROR:
    MsgBox Err.Description, vbInformation, "Aviso"
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As ADODB.Recordset
Dim oAge As New DActualizaDatosArea

txtIngCaja = gdFecSis
CentraForm Me
lstAgencias.Clear
Set rs = oAge.GetAgencias(, False)
If RSVacio(rs) Then
Else
   Do While Not rs.EOF
     Me.lstAgencias.AddItem (rs!Descripcion & Space(100) & Trim(rs!Codigo))
     rs.MoveNext
   Loop
End If
Me.lvCheque.ColumnHeaders(3).Width = 0
RSClose rs
End Sub

Private Sub lstAgencias_Click()
Dim lsAgencia As String
End Sub

Private Sub lvCheque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvCheque.SortKey = ColumnHeader.SubItemIndex
    If lvCheque.SortOrder = lvwAscending Then
        lvCheque.SortOrder = lvwDescending
    Else
        lvCheque.SortOrder = lvwAscending
    End If
    lvCheque.Sorted = True
End Sub

Private Sub lvCheque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
   lnMonto = lnMonto + Item.SubItems(4)
Else
   lnMonto = lnMonto - Item.SubItems(4)
End If
txtMonto.Text = Format(lnMonto, "#,#00.00")
End Sub

'Private Sub txtAgeCod_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then
'   txtAgeDesc.Text = DameTablaCod("47", Trim(txtAgeCod))
'   Call MuestraCheques(txtAgeCod)
' End If
'End Sub
