VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCierreContAnual 
   Caption         =   "Cierre Contable Anual"
   ClientHeight    =   5460
   ClientLeft      =   2325
   ClientTop       =   2280
   ClientWidth     =   5430
   Icon            =   "frmCierreContAnual.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      Picture         =   "frmCierreContAnual.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   4815
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   1755
      TabIndex        =   0
      Top             =   3555
      Width           =   2175
      Begin VB.TextBox txtAnio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   945
         MaxLength       =   4
         TabIndex        =   2
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AÑO"
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
         Left            =   270
         TabIndex        =   1
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5190
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   5130
      Begin VB.Frame Frame3 
         Height          =   1530
         Left            =   90
         TabIndex        =   11
         Top             =   1440
         Width           =   4965
         Begin VB.PictureBox pic4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            Picture         =   "frmCierreContAnual.frx":064C
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   19
            Top             =   1140
            Width           =   255
         End
         Begin VB.PictureBox pic3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            Picture         =   "frmCierreContAnual.frx":098E
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   18
            Top             =   820
            Width           =   255
         End
         Begin VB.PictureBox pic2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            Picture         =   "frmCierreContAnual.frx":0CD0
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   17
            Top             =   500
            Width           =   255
         End
         Begin VB.PictureBox pic1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            Picture         =   "frmCierreContAnual.frx":1012
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   16
            Top             =   180
            Width           =   255
         End
         Begin VB.Label lbl4 
            Caption         =   "Resultado Neto a Utilidad/Perdida Neta"
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
            Height          =   240
            Left            =   435
            TabIndex        =   20
            Top             =   1170
            Width           =   4335
         End
         Begin VB.Label lbl3 
            Caption         =   "Resultados de Ejercicio a Resultado Neto"
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
            Height          =   240
            Left            =   435
            TabIndex        =   14
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label lbl2 
            Caption         =   "Resultados de Operación a Resultados de Ejercicio"
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
            Height          =   240
            Left            =   435
            TabIndex        =   13
            Top             =   495
            Width           =   4470
         End
         Begin VB.Label lbl1 
            Caption         =   "Ingresos y Gastos a Resultados de Operación"
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
            Left            =   435
            TabIndex        =   12
            Top             =   210
            Width           =   4305
         End
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Cierre Contable"
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
         Height          =   1125
         Left            =   1440
         TabIndex        =   7
         Top             =   3255
         Width           =   2505
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
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
         Left            =   2730
         TabIndex        =   6
         Top             =   4545
         Width           =   1245
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Default         =   -1  'True
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
         Left            =   1380
         TabIndex        =   5
         Top             =   4545
         Width           =   1245
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   1185
         Left            =   120
         TabIndex        =   4
         Top             =   3285
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2090
         _Version        =   393216
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         Enabled         =   0   'False
         FullWidth       =   85
         FullHeight      =   79
      End
      Begin VB.Image imgAlerta 
         Height          =   480
         Left            =   780
         Picture         =   "frmCierreContAnual.frx":1354
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ADVERTENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1770
         TabIndex        =   9
         Top             =   330
         Width           =   1905
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Este proceso realiza el Cierre Anual del Ejercicio. El Sistema no permitirá el Ingreso de Movimientos en un Ejercicio Cerrado."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   630
         TabIndex        =   8
         Top             =   690
         Width           =   4020
         WordWrap        =   -1  'True
      End
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   270
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmCierreContAnual.frx":14A5
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
Attribute VB_Name = "frmCierreContAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs   As ADODB.Recordset
Dim nTot  As Currency
Dim nTot6 As Currency
Dim sCtaRes As String
Dim sCtaResPerdida As String
Dim dFecCierre As Date
Dim oMov As DMov

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Function Valida() As Boolean
Valida = True
If txtAnio = "" Then
   MsgBox "Debe ingresar Año de Cierre!", vbInformation, "!Aviso!"
   Exit Function
End If
If Val(txtAnio) > Year(gdFecSis) Then
   MsgBox "Año de Cierre mayor que año actual!", vbInformation, "!Aviso!"
   Exit Function
End If
If Year(gdFecSis) - Val(txtAnio) > 2 Then
   MsgBox "Imposible realizar Cierre de periodos anteriores", vbInformation, "!Aviso!"
   Exit Function
End If
End Function

Private Sub Controles(lActiva As Boolean)
cmdProcesar.Enabled = lActiva
cmdCancelar.Enabled = lActiva
End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdProcesar_Click()
Dim sFec As String
Dim sAge As String
Dim sTpo As String
Dim sClave As String
Dim n As Currency


On Error GoTo ErrorCierreCont
If Not Valida Then Exit Sub
   dFecCierre = CDate("01/01/" & Val(txtAnio) + 1) - 1

   If MsgBox("¿ Desea Realizar Cierre Anual de Contabilidad ?", vbYesNo + vbQuestion, "¡Confirmación!") = vbNo Then
      Exit Sub
   End If

   gsMovNro = Format(dFecCierre, gsFormatoMovFecha) + "235959" + gsCodCMAC
   Dim oCont As New NContFunciones
   If oCont.ExisteMovimiento(Left(gsMovNro, 8), Left(gContCierreAnual, 5)) Then
       MsgBox "Cierre Anual del Ejercicio " & txtAnio & " ya fue Realizado", vbInformation, "Aviso"
       cmdProcesar.Visible = True
       cmdCancelar.Visible = True
       Screen.MousePointer = 0
       Set oCont = Nothing
       Exit Sub
   End If
   Set oCont = Nothing
   gsMovNro = Format(dFecCierre, gsFormatoMovFecha) + "235959" + gsCodCMAC & Right(gsCodAge, 2) + "00" & gsCodUser
   
   Controles False
   sFec = dFecCierre
   nTot = 0
   nTot6 = 0
   rtxt.Text = ""
   sCtaRes = ""
   lbl1.ForeColor = vbRed
   Dim lbSoloAjustadas As Boolean
   Set oMov = New DMov
   If gsCodCMAC = "102" Then  'Caja de Lima
      lbSoloAjustadas = True
   Else
      lbSoloAjustadas = False
   End If
   
    rtxt.Text = rtxt.Text & ProcesaAsiento(1, "", 0, 0, lbSoloAjustadas) & oImpresora.gPrnSaltoPagina
    If lbSoloAjustadas Then
       rtxt.Text = rtxt.Text & ProcesaAsientoAjustadas(5) & oImpresora.gPrnSaltoPagina
    End If
    pic1.Picture = picX.Picture
    
    lbl2.ForeColor = vbRed
    rtxt.Text = rtxt.Text & ProcesaAsiento(2, sCtaRes, nTot, nTot6, lbSoloAjustadas) & oImpresora.gPrnSaltoPagina
    pic2.Picture = picX.Picture
     
    rtxt.Text = rtxt.Text & ProcesaAsiento(3, sCtaRes, nTot, nTot6, lbSoloAjustadas) & oImpresora.gPrnSaltoPagina
    pic3.Picture = picX.Picture
    lbl3.ForeColor = vbRed
   
   If gsCodCMAC = "109" Then  'Malditos trujillanos que no cambiaron el codigo de cmac GITU
       rtxt.Text = rtxt.Text & ProcesaAsiento(4, sCtaRes, nTot, nTot6, lbSoloAjustadas) & oImpresora.gPrnSaltoPagina
       pic4.Picture = picX.Picture
       lbl4.ForeColor = vbRed
   End If
   Set oMov = Nothing
   
   Dim oGen As New NConstSistemas
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Set oGen = Nothing
   Screen.MousePointer = 0
   EnviaPrevio rtxt.Text, "Asientos de Cierre Anual", gnLinPage, False
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Cierre de Año Realizado Satisfactoriamente con Fecha de Cierre " & dFecCierre
            Set objPista = Nothing
            '*******
   Unload Me
Exit Sub
ErrorCierreCont:
    MsgBox "Error Nº [" & Err.Number & "] " & TextErr(Err.Description) & Chr(13) & "Consulte al Area de Sistemas", vbInformation, "Aviso"

End Sub

Private Function ProcesaAsiento(nTipo As Integer, ByVal psCtaRes As String, ByVal pnTot As Currency, ByVal pnTot6 As Currency, ByVal pbSoloAjustadas As Boolean) As String
Dim sCond   As String
Dim sCond1  As String
Dim sCond2  As String

Dim nItem As Integer
Dim nImporte As Currency
Dim sOpeCod  As String
Dim sItem As String
Dim oOpe  As New DOperacion
On Error GoTo ErrProcesa
sOpeCod = Format(Val(gsOpeCod) + nTipo, "000000")

   Set rs = oOpe.CargaOpeCta(sOpeCod, , "0")
   If Not rs.EOF Then
      Do While Not rs.EOF
         If Len(rs!cCtaContCod) = 1 Then
            sCond1 = sCond1 & "'" & rs!cCtaContCod & "',"
         Else
            sCond2 = sCond2 & "'" & rs!cCtaContCod & "',"
         End If
         rs.MoveNext
      Loop
      If sCond1 <> "" Then
         sCond1 = Mid(sCond1, 1, Len(sCond1) - 1)
      End If
      If sCond2 <> "" Then
         sCond2 = Mid(sCond2, 1, Len(sCond2) - 1)
      End If
   Else
      MsgBox "No se definieron las cuentas a procesar en la Operación", vbInformation, "Aviso"
      Exit Function
   End If
   
'Nuevo Cuenta del Haber
   Set rs = oOpe.CargaOpeCta(sOpeCod, , "1")
   If Not rs.EOF Then
      sCtaRes = rs!cCtaContCod
   End If
      
   If nTipo = 4 Then   'Traslado a la Utilidad o Perdida
      Set rs = oOpe.CargaOpeCta(sOpeCod, , "2")
      If Not rs.EOF Then
         sCtaResPerdida = rs!cCtaContCod
      End If
   End If
   If sCond1 <> "" And sCond2 <> "" Then
      sCond = "(substring(c.cCtaContCod,1,1) IN (" & sCond1 & ") " _
           & " or substring(c.cCtaContCod,1,2) IN (" & sCond2 & ") )"
   ElseIf sCond1 <> "" And sCond2 = "" Then
      sCond = "substring(c.cCtaContCod,1,1) IN (" & sCond1 & ") "
   Else
      sCond = "substring(c.cCtaContCod,1,2) IN (" & sCond2 & ") "
   End If
   Dim oAjuste As New DAjusteCont
   Set rs = oAjuste.VerAjusteCierreAnual(Format(CDate(GetFechaMov(gsMovNro, True)), gsFormatoFecha), sCond)
   Set oAjuste = Nothing
'   If rs.EOF And rs.BOF Then
'      MsgBox "No existen Cuentas para Traslado...", vbInformation, "Advertencia"
'      RSClose rs
'      cmdProcesar.Visible = True
'      cmdCancelar.Visible = True
'      Screen.MousePointer = 0
'      Exit Function
'   End If
   
   gsMovNro = oMov.GeneraMovNro(gdFecSis, , , gsMovNro)
   
   Select Case nTipo
      Case 1: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Ingresos/Egresos a Resultados de Operación"
      Case 2: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Resultados de Operación a Resultados del Ejercicio"
      Case 3: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Resultados del Ejercicio a Resultado Neto"
   End Select
   If nTipo = 4 Then
      gsGlosa = "Cierre Anual " + txtAnio & ": Resultado Neto a " & IIf(pnTot6 > 0, "Pérdida", "Utilidad") & " del Ejercicio "
   End If
   oMov.BeginTrans
   oMov.InsertaMov gsMovNro, sOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
   gnMovNro = oMov.GetnMovNro(gsMovNro)
   oMov.InsertaMovCont gnMovNro, 0, 0, ""
   oMov.InsertaMovOtrosItem gnMovNro, 1, "CieAnio" & nTipo, Format(txtAnio, "0000"), ""
   
   nItem = 0
   nTot = 0: nTot6 = 0
   If Not rs.EOF Then
        Do While Not rs.EOF
           DoEvents
           nImporte = 0
           If psCtaRes <> "" Then
              If Mid(rs!cCtaContCod, 3, 1) = 6 Then
                 If Mid(psCtaRes, 1, 2) & "6" & Mid(psCtaRes, 4) = rs!cCtaContCod Then
                    nImporte = pnTot6
                 End If
              Else
                 If psCtaRes = rs!cCtaContCod Then
                    nImporte = pnTot
                 End If
              End If
           End If
           nImporte = nImporte + rs!nCtaSaldoImporte
              
           If nImporte <> 0 Then
              nItem = nItem + 1
              If Mid(rs!cTipo, 1, 1) = "D" Then 'Deudora
                 If (Mid(rs!cCtaContCod, 3, 1) = "6" And Not pbSoloAjustadas) Or Not Mid(rs!cCtaContCod, 3, 1) = "6" Then
                    oMov.InsertaMovCta gnMovNro, nItem, rs!cCtaContCod, nImporte * -1
                 End If
                 If Mid(rs!cCtaContCod, 3, 1) = "6" Then
                    nTot6 = nTot6 + (nImporte)
                 Else
                    nTot = nTot + (nImporte)
                 End If
              Else 'Acreedora
                 If (Mid(rs!cCtaContCod, 3, 1) = "6" And Not pbSoloAjustadas) Or Not Mid(rs!cCtaContCod, 3, 1) = "6" Then
                     oMov.InsertaMovCta gnMovNro, nItem, rs!cCtaContCod, nImporte
                 End If
                 If Mid(rs!cCtaContCod, 3, 1) = "6" Then
                    nTot6 = nTot6 + (nImporte * -1)
                 Else
                    nTot = nTot + (nImporte * -1)
                 End If
              End If
           End If
           rs.MoveNext
        Loop
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, sCtaRes, nTot
   Else
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, psCtaRes, pnTot * -1
        nTot = pnTot
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, sCtaRes, pnTot
   End If
   RSClose rs
   If nTot6 <> 0 And Not nTipo = 4 And Not pbSoloAjustadas Then
       nItem = nItem + 1
       oMov.InsertaMovCta gnMovNro, nItem, Mid(IIf(pnTot < 0 And sCtaResPerdida <> "", sCtaResPerdida, sCtaRes), 1, 2) & "6" & Mid(IIf(pnTot < 0 And sCtaResPerdida <> "", sCtaResPerdida, sCtaRes), 4), nTot6
   End If
   'ALPA 20120418***************************************
   oMov.GeneraMovMECierreAnual gnMovNro, gsMovNro
   '****************************************************
   oMov.CommitTrans

   sOpeCod = gsOpeCod
   Dim oCont As New NContImprimir
   ProcesaAsiento = oCont.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "Cierre Anual: Asiento Contable")
   Set oCont = Nothing
   gsOpeCod = sOpeCod
Exit Function
ErrProcesa:
    oMov.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ProcesaAsientoAjustadas(nTipo As Integer) As String
Dim sCond   As String
Dim sCond1  As String
Dim sCond2  As String

Dim nItem As Integer
Dim nImporte As Currency
Dim sOpeCod  As String
Dim sItem As String
Dim oOpe  As New DOperacion
Dim lsCtaGasto As String
Dim lsCtaIngre As String
Dim lsCtaResul As String
Dim nTotG As Currency
Dim nTotI As Currency
Dim nTotR As Currency

On Error GoTo ErrProcesa
sOpeCod = Format(Val(gsOpeCod) + nTipo, "000000")

   Set rs = oOpe.CargaOpeCta(sOpeCod, , "0")
   If Not rs.EOF Then
      Do While Not rs.EOF
         If Len(rs!cCtaContCod) = 1 Then
            sCond1 = sCond1 & "'" & rs!cCtaContCod & "',"
         Else
            sCond2 = sCond2 & "'" & rs!cCtaContCod & "',"
         End If
         rs.MoveNext
      Loop
      If sCond1 <> "" Then
         sCond1 = Mid(sCond1, 1, Len(sCond1) - 1)
      End If
      If sCond2 <> "" Then
         sCond2 = Mid(sCond2, 1, Len(sCond2) - 1)
      End If
   Else
      MsgBox "No se definieron las cuentas a procesar en la Operación", vbInformation, "Aviso"
      Exit Function
   End If
   
'Nuevo Cuenta del Haber
   Set rs = oOpe.CargaOpeCta(sOpeCod, , "1")
   If Not rs.EOF Then
      lsCtaGasto = rs!cCtaContCod
   End If
   Set rs = oOpe.CargaOpeCta(sOpeCod, , "2")
   If Not rs.EOF Then
      lsCtaIngre = rs!cCtaContCod
   End If
   Set rs = oOpe.CargaOpeCta(sOpeCod, , "3")
   If Not rs.EOF Then
      lsCtaResul = rs!cCtaContCod
   End If
   
   If sCond1 <> "" And sCond2 <> "" Then
      sCond = "(substring(c.cCtaContCod,1,1) IN (" & sCond1 & ") " _
           & " or substring(c.cCtaContCod,1,2) IN (" & sCond2 & ") ) and SubString(c.cCtaContCod,3,1) = '6' "
   ElseIf sCond1 <> "" And sCond2 = "" Then
      sCond = "substring(c.cCtaContCod,1,1) IN (" & sCond1 & ")  and SubString(c.cCtaContCod,3,1) = '6' "
   Else
      sCond = "substring(c.cCtaContCod,1,2) IN (" & sCond2 & ")  and SubString(c.cCtaContCod,3,1) = '6' "
   End If
   Dim oAjuste As New DAjusteCont
   Set rs = oAjuste.VerAjusteCierreAnual(Format(CDate(GetFechaMov(gsMovNro, True)), gsFormatoFecha), sCond)
   Set oAjuste = Nothing
'   If rs.EOF And rs.BOF Then
'      MsgBox "No existen Cuentas para Traslado...", vbInformation, "Advertencia"
'      RSClose rs
'      cmdProcesar.Visible = True
'      cmdCancelar.Visible = True
'      Screen.MousePointer = 0
'      Exit Function
'   End If
   
   gsMovNro = oMov.GeneraMovNro(gdFecSis, , , gsMovNro)
   
   Select Case nTipo
      Case 1: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Ingresos/Egresos a Resultados de Operación"
      Case 2: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Resultados de Operación a Resultados del Ejercicio"
      Case 3: gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Resultados del Ejercicio a Resultado Neto"
   End Select
   gsGlosa = "Cierre Anual " + txtAnio & ": Saldos de Cuentas Ajustadas a Cuentas de Resultado"
   oMov.BeginTrans
   oMov.InsertaMov gsMovNro, sOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
   gnMovNro = oMov.GetnMovNro(gsMovNro)
   oMov.InsertaMovCont gnMovNro, 0, 0, ""
   oMov.InsertaMovOtrosItem gnMovNro, 1, "CieAnio" & nTipo, Format(txtAnio, "0000"), ""
   
   nItem = 0
   nTotG = 0: nTotI = 0: nTotR = 0
   If Not rs.EOF Then
        Do While Not rs.EOF
           DoEvents
           nImporte = rs!nCtaSaldoImporte
              
           If nImporte <> 0 Then
              nItem = nItem + 1
              If Mid(rs!cTipo, 1, 1) = "D" Then 'Deudora
                 oMov.InsertaMovCta gnMovNro, nItem, rs!cCtaContCod, nImporte * -1
                 If Mid(rs!cCtaContCod, 1, 1) = "4" Then
                    nTotG = nTotG + (nImporte)
                 ElseIf Mid(rs!cCtaContCod, 1, 1) = "5" Then
                    nTotI = nTotI + (nImporte)
                 Else
                    nTotR = nTotR + (nImporte)
                 End If
              Else 'Acreedora
                 oMov.InsertaMovCta gnMovNro, nItem, rs!cCtaContCod, nImporte
                 If Mid(rs!cCtaContCod, 1, 1) = "4" Then
                    nTotG = nTotG + (nImporte * -1)
                 ElseIf Mid(rs!cCtaContCod, 1, 1) = "5" Then
                    nTotI = nTotI + (nImporte * -1)
                 Else
                    nTotR = nTotR + (nImporte * -1)
                 End If
              End If
           End If
           rs.MoveNext
        Loop
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, lsCtaGasto, nTotG
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, lsCtaIngre, nTotI
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, lsCtaResul, nTotR
   End If
   RSClose rs
   oMov.CommitTrans

   sOpeCod = gsOpeCod
   Dim oCont As New NContImprimir
   ProcesaAsientoAjustadas = oCont.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "Cierre Anual: Asiento Contable")
   Set oCont = Nothing
   gsOpeCod = sOpeCod
Exit Function
ErrProcesa:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Form_Load()
Dim sUltCierre As String
CentraForm Me
sUltCierre = LeeConstanteSist(gConstSistCierreAnualCont)
txtAnio = Val(sUltCierre) + 1

If gsCodCMAC = "112" Then
    pic4.Visible = False
    lbl4.Visible = False
End If
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cmdProcesar.SetFocus
End If
End Sub
