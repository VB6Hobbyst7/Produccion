VERSION 5.00
Begin VB.Form frmCredExonerarPreCanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneración Comisión PreCancelación"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmCredExonerarPreCanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   325
      Left            =   3840
      TabIndex        =   19
      ToolTipText     =   "Busca cliente por nombre, documento o codigo"
      Top             =   150
      Width           =   375
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   4320
      TabIndex        =   18
      Top             =   3120
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1050
   End
   Begin VB.CommandButton cmdExonerar 
      Caption         =   "Exonerar"
      Height          =   360
      Left            =   3120
      TabIndex        =   16
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   1900
         Width           =   4395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1940
         Width           =   495
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Comisión :"
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   1485
         Width           =   720
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   640
         Width           =   4140
      End
      Begin VB.Label lblTitular 
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   285
         Width           =   4140
      End
      Begin VB.Label lblSaldoCanc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Cancelación :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1480
         Width           =   1425
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo :"
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   1110
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   640
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   285
         Width           =   525
      End
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   847
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "109"
   End
End
Attribute VB_Name = "frmCredExonerarPreCanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredAdmExoComPreCanc
'** Descripción : Formulario para exonerar comisión para PreCancelaciones según Anexo 02 TI-ERS097-2013
'** Creación : JUEZ, 20130923 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaDatos(ActxCta.NroCuenta)
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, ActxCta)
        ActxCta.SetFocusCuenta
    Else
        cmdCancelar_Click
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    LimpiaPantalla
End Sub

Private Sub cmdExonerar_Click()
    If Trim(txtGlosa.Text) = "" Then
        MsgBox "Es necesario ingresar el motivo de la exoneración", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se va a exonerar el pago de comisión por precancelación al crédito. Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim sMovNro As String
    Dim oDCred As COMDCredito.DCOMCredActBD
    Set oDCred = New COMDCredito.DCOMCredActBD
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oDCred.dInsertExonerarPreCanc(Me.ActxCta.NroCuenta, gdFecSis, CDbl(lblSaldoCanc.Caption), CDbl(lblComision.Caption), Trim(txtGlosa.Text), sMovNro)
    MsgBox "La exoneración fue registrada con éxito. Recuerde que es valida sólo en el día en que se emite.", vbInformation, "Aviso"
    cmdCancelar_Click
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ActxCta.Age = gsCodAge
    cmdExonerar.Enabled = False
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oCredito As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim rs As ADODB.Recordset, rsCredVig As ADODB.Recordset
    Set oCredito = New COMDCredito.DCOMCredito
    Set rs = oCredito.RecuperaDatosComunes(psCtaCod, False)
    If rs.EOF Or rs.BOF Then
        MsgBox "Credito No Existe o No Esta Vigente", vbInformation, "Aviso"
        LimpiaPantalla
        Exit Sub
    End If
    If rs.RecordCount > 0 Then
        If rs!nPersPersoneria = 1 Then  'Comisión sólo para Pers Jur
            MsgBox "El titular del crédito debe ser una Persona Juridica", vbInformation, "Aviso"
            LimpiaPantalla
            Exit Sub
        End If
        If (Left(rs!cTpoCredCod, 1) <> "1" And Left(rs!cTpoCredCod, 1) <> "2" And Left(rs!cTpoCredCod, 1) <> "3") Then 'Comisión sólo para Tipo Cred Corporativa, Grande o Mediana Empresa
            MsgBox "El tipo de crédito debe ser Mediana Empresa, Grande Empresa o Corporativo", vbInformation, "Aviso"
            LimpiaPantalla
            Exit Sub
        End If
        If Not oCredito.VerificaSiEsCancelacionAnticipada(psCtaCod, gdFecSis) Then
            MsgBox "El crédito no tiene más de 2 cuotas pendientes", vbInformation, "Aviso"
            LimpiaPantalla
            Exit Sub
        End If
        lblTitular.Caption = PstaNombre(rs!cTitular)
        LblAnalista.Caption = PstaNombre(rs!cAnalista)
        LblPrestamo.Caption = Format(rs!nMontoCol, "#,##0.00")
        LblMoneda.Caption = rs!cMoneda
        
        Dim pMatCalend As Variant
        Dim nInteresFecha As Currency
        Dim nInterFechaGra As Currency
        Dim nMontoFecha As Currency
        Dim nSaldoCanc As Currency
        
        Set rsCredVig = oCredito.RecuperaDatosCreditoVigente(psCtaCod, gdFecSis)
        
        Set oNCred = New COMNCredito.NCOMCredito
        pMatCalend = oNCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
        nInteresFecha = oNCred.MatrizInteresGastosAFecha(psCtaCod, pMatCalend, gdFecSis, True, IIf(IsNull(rsCredVig!nCalendDinamico), False, IIf(rsCredVig!nCalendDinamico = 0, False, True)))
        nInterFechaGra = oNCred.MatrizInteresGraAFecha(psCtaCod, pMatCalend, gdFecSis, True, IIf(IsNull(rsCredVig!nCalendDinamico), False, IIf(rsCredVig!nCalendDinamico = 0, False, True)))
        nMontoFecha = oNCred.MatrizCapitalAFecha(psCtaCod, pMatCalend)
        Set oNCred = Nothing
        
        nSaldoCanc = Format(nInteresFecha + nInterFechaGra + nMontoFecha, "#0.00")
        lblSaldoCanc.Caption = Format(nSaldoCanc, "#,##0.00")
        
        lblComision.Caption = Format(CalculaComisionPreCancelacion(nSaldoCanc, psCtaCod), "#,##0.00")
        ActxCta.Enabled = False
        cmdBuscar.Enabled = False
        cmdExonerar.Enabled = True
    Else
        MsgBox "Credito No Existe o No Esta Vigente", vbInformation, "Aviso"
        LimpiaPantalla
    End If
    rs.Close
    Set oCredito = Nothing
End Sub

Private Sub LimpiaPantalla()
    ActxCta.Enabled = True
    cmdBuscar.Enabled = True
    ActxCta.NroCuenta = ""
    ActxCta.Prod = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblTitular.Caption = ""
    LblAnalista.Caption = ""
    LblPrestamo.Caption = ""
    lblSaldoCanc.Caption = ""
    LblMoneda.Caption = ""
    lblComision.Caption = ""
    txtGlosa.Text = ""
    cmdExonerar.Enabled = False
    ActxCta.SetFocus
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExonerar.SetFocus
    End If
End Sub
