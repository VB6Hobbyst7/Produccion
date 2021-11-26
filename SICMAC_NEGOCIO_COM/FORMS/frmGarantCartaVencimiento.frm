VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGarantCartaVencimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta de Vencimiento de Tasación"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmGarantCartaVencimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   310
      Left            =   2280
      TabIndex        =   10
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Frame FrameAgencias 
      Caption         =   " Agencias "
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   3315
      Begin VB.OptionButton OptAgencias 
         Caption         =   "Ag. Local"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptAgencias 
         Caption         =   "Ag. Remotas"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   310
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Frame FramePeriodo 
      Caption         =   " Período "
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3315
      Begin MSMask.MaskEdBox mskFechaCierre 
         Height          =   315
         Left            =   1390
         TabIndex        =   0
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
      Begin VB.Label Label1 
         Caption         =   "Fecha Cierre:"
         Height          =   255
         Left            =   340
         TabIndex        =   15
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame FrameTipo 
      Caption         =   " Tipo "
      ForeColor       =   &H00FF0000&
      Height          =   1770
      Left            =   120
      TabIndex        =   13
      Top             =   1140
      Width           =   3315
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Corporativos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Grandes empresas"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Width           =   1695
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Medianas empresas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   810
         Width           =   1995
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Pequeñas empresas"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1095
         Width           =   1875
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Microempresas"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1380
         Width           =   1695
      End
   End
   Begin VB.Frame FrameMoneda 
      Caption         =   " Moneda "
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   120
      TabIndex        =   12
      Top             =   2920
      Width           =   3315
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Nacional"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Extranjera"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGarantCartaVencimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'CREADOR:       RECO, Renzo Córdova
'FECHA:         08/01/2016
'DESCRIPCION:   Formulario que permite la generación de cartas automaticas de vencimiento de tasacion para creditos con garantia real
'************************************************************
Option Explicit
Dim Agencias As String

Private Function ObtieneMoneda() As String
    If ChkMoneda(0).value = 1 Then ObtieneMoneda = "1"
    If ChkMoneda(1).value = 1 Then ObtieneMoneda = "2"
    If ChkMoneda(0).value = 1 And ChkMoneda(1).value = 1 Then ObtieneMoneda = "1,2"
End Function
Private Function ObtieneTipoCred() As String
    If ChkTipo(0).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "1,"
    If ChkTipo(1).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "2,"
    If ChkTipo(2).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "3,"
    If ChkTipo(3).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "4,"
    If ChkTipo(4).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "5,"
    ObtieneTipoCred = Mid(ObtieneTipoCred, 1, Len(ObtieneTipoCred) - 1)
End Function

Private Sub GeneraCartaAvisoVencimiento(ByVal prsDatosCartas As ADODB.Recordset)
    On Error GoTo ErrorImprimirPDF
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim oDoc  As New cPDF
    Dim nTipo As Integer, nPosicion As Integer, nIndice As Integer
    Dim sNroCarta As String
    Dim nIndexChar As Integer
    Dim nTop As Integer
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Carta de Aviso por Vencimiento de Tasacion" & "RECO"
    oDoc.Title = "Carta de Aviso por Vencimiento de Tasacion" & "RECO"
    
    
    If Not (prsDatosCartas.EOF And prsDatosCartas.BOF) Then
        
        If Not oDoc.PDFCreate(App.path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & "GarantReal_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
            Exit Sub
        End If
    
    
            oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding
            oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding
            oDoc.Fonts.Add "F3", "Arial", TrueType, Normal, WinAnsiEncoding
            oDoc.Fonts.Add "F4", "Arial", TrueType, BoldItalic, WinAnsiEncoding
            oDoc.Fonts.Add "F5", "Arial", TrueType, Italic, WinAnsiEncoding
            oDoc.Fonts.Add "F", "Times New Roman", TrueType, Bold, WinAnsiEncoding
        For nIndice = 1 To prsDatosCartas.RecordCount
    
            sNroCarta = oGarant.GarantRealRegistraCartaVenc(prsDatosCartas!Credito, gsCodUser)
    
            Dim nFTabla As Integer
            Dim nFTablaCabecera As Integer
            Dim lnFontSizeBody As Integer
        
            nFTablaCabecera = 7
            nFTabla = 7
            lnFontSizeBody = 7
            nTop = 50
            nIndexChar = InStr(prsDatosCartas!cUbiGeoDescripcion, "(")
            oDoc.NewPage A4_Vertical
            oDoc.WTextBox 40 + nTop, 70, 10, 450, "", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 40 + nTop, 70, 10, 450, "CARTA Nº " & sNroCarta & "-TV-SO-CMAC-M.", "F4", 11, hCenter, , vbBlack
            oDoc.WTextBox 80 + nTop, 70, 10, 450, Mid(prsDatosCartas!cUbiGeoDescripcion, 1, 1) & LCase(Mid(prsDatosCartas!cUbiGeoDescripcion, 2, nIndexChar - 2)) & ", " & IIf(Len(Day(Date)) = 1, "0" & Day(Date), Day(Date)) & " de " & fgDameNombreMes(Month(Date)) & " del " & Year(Date), "F3", 11, hRight, , vbBlack
            
            oDoc.WTextBox 115 + nTop, 70, 10, 450, "Señor (Es):", "F4", 11, hLeft, , vbBlack
            oDoc.WTextBox 130 + nTop, 70, 10, 450, prsDatosCartas!Cliente, "F4", 11, hLeft, , vbBlack
            oDoc.WTextBox 165 + nTop, 70, 10, 450, "Dirc: " & prsDatosCartas!cPersDireccDomicilio, "F4", 11, hLeft, , vbBlack
            oDoc.WTextBox 180 + nTop, 70, 10, 450, "Presente.-", "F3", 11, hLeft, , vbBlack
            
            oDoc.WTextBox 200 + nTop, 70, 10, 450, "Es grato dirigirme a usted, con la finalidad de saludarlo y al mismo tiempo comunicarle  que:", "F3", 11, hjustify, , vbBlack
            
            oDoc.WTextBox 225 + nTop, 70, 10, 450, "“…De conformidad con lo dispuesto en la circular B-2184-2010- emitida por la Superintendencia de Banca y Seguros, LA " & _
                                            "PRESTATARIA/GARANTE se obliga a remitir la información actualizada cada vez que sea requerido por LA CAJA cuyo cumplimiento " & _
                                            "deberá ejecutarse a un plazo no mayor de 3 días a partir de la fecha de recepción del requerimiento.”", "F4", 11, hjustify, , vbBlack
                                            
            oDoc.WTextBox 280 + nTop, 70, 10, 450, "Asimismo en el Contrato de Mutuo con Garantía Hipotecaria firmado con nuestra entidad, en su  cláusula cuarta establece  lo siguiente:", "F3", 11, hjustify, , vbBlack
            
            oDoc.WTextBox 315 + nTop, 70, 10, 450, "“…En caso LA CAJA requiera la renovación de la tasación del bien hipotecado EL PRESTATARIO y/o  LOS GARANTES HIPOTECARIOS se obligan " & _
                                            "a remitir la tasación actualizada (perito tasador autorizado por LA CAJA), cuyo plazo debe ejecutarse en un plazo no mayor de 3 días a " & _
                                            "partir de la fecha de recepción, del requerimiento.", "F5", 11, hjustify, , vbBlack
                                            
            oDoc.WTextBox 370 + nTop, 70, 10, 450, " En caso que EL PRESTATARIO yo GARANTES HIPOTECARIOS  no cumplieran con lo requerido, LA CAJA queda autorizada a contratar por cuenta, " & _
                                            "costo y riesgo de EL PRESTATARIO a un perito tasador autorizado por LA CAJA para la actualización de la tasación, costo que será cargado " & _
                                            "en la cuota de vencimiento mas próxima, del calendario de pago pactado o del crédito vencido”;", "F5", 11, hjustify, , vbBlack
                                            
            oDoc.WTextBox 445 + nTop, 70, 10, 450, "por lo que solicitamos alcanzar:", "F3", 11, hjustify, , vbBlack
            
            oDoc.WTextBox 460 + nTop, 130, 10, 450, "•   TASACION  ACTUALIZADA DEL INMUEBLE ubicado en:", "F3", 11, hjustify, , vbBlack
            oDoc.WTextBox 475 + nTop, 130, 10, 450, prsDatosCartas!cDireccion, "F4", 11, hjustify, , vbBlack
            oDoc.WTextBox 515 + nTop, 70, 10, 450, "Por lo expuesto y de no alcanzar lo solicitado en el plazo estipulado, procederemos a contratar al perito tasador   realizar el cargo del costo en las cuotas a vencer según el cronograma de pagos. Ello en cumplimiento de lo establecido en la cláusula antes mencionada", "F3", 11, hjustify, , vbBlack
            oDoc.WTextBox 595 + nTop, 70, 10, 450, "Atentamente,", "F3", 11, hjustify, , vbBlack

            
            
            prsDatosCartas.MoveNext
        Next
    End If

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox err.Description, vbInformation, "Aviso"

End Sub

Private Sub cmdImprimir_Click()
    Screen.MousePointer = 11
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim sMsj As String
    sMsj = ValidaDatos
    If sMsj = "" Then
        Agencias = IIf(OptAgencias(0).value = True, gsCodAge, frmSelectAgencias.RecupAgencias)
        Agencias = Replace(Replace(Agencias, "'", ""), " ", "")
        Agencias = Replace(Agencias, "(", "")
        Agencias = Replace(Agencias, ")", "")
        Call GeneraCartaAvisoVencimiento(oGarant.ObtieneDatosGarantRealVenc(Format(mskFechaCierre.Text, "yyyy/MM/dd"), Agencias, ObtieneTipoCred, ObtieneMoneda, "2020,2021,2022,2030,2031,2032,2201,2202"))
    Else
        MsgBox sMsj, vbInformation, "Alerta"
    End If
    Screen.MousePointer = 0
End Sub

Private Function DameAgencias() As String
    Dim Agencias As String
    Dim lnAge As Integer
    Dim est As Integer
    est = 0
    Agencias = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            est = est + 1
            If est = 1 Then
                Agencias = "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
            Else
                Agencias = Agencias & ", " & "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
            End If
        End If
    Next lnAge
    DameAgencias = Agencias
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub mskFechaCierre_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(mskFechaCierre.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(mskFechaCierre.Text) > gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            mskFechaCierre.SetFocus
            Exit Sub
        End If
End Sub

Private Sub OptAgencias_Click(index As Integer)
    If index = 1 Then
        frmSelectAgencias.Inicio Me
        frmSelectAgencias.Show vbModal
    End If
End Sub

Private Function ValidaDatos() As String
    ValidaDatos = ""
    If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 Then
        ValidaDatos = "Debe seleccionar por lo menos un tipo de crédito"
        Exit Function
    End If
    If ChkMoneda(0).value = 0 And ChkMoneda(1).value = 0 Then
        ValidaDatos = "Debe seleccionar por lo menos un tipo de moneda"
        Exit Function
    End If
    If mskFechaCierre.Text = "__/__/____" Then
        ValidaDatos = "Debe ingresar la fecha"
        Exit Function
    End If
    
    ValidaDatos = ValidaFecha(mskFechaCierre.Text)
    If ValidaDatos <> "" Then
        Exit Function
    End If
    If CDate(mskFechaCierre.Text) > gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        mskFechaCierre.SetFocus
        Exit Function
    End If
End Function
