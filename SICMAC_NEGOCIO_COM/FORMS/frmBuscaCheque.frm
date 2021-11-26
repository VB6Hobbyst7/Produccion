VERSION 5.00
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmBuscaCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscando Cheque"
   ClientHeight    =   2955
   ClientLeft      =   4530
   ClientTop       =   3810
   ClientWidth     =   4710
   Icon            =   "frmBuscaCheque.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3330
      TabIndex        =   10
      Top             =   2520
      Width           =   1140
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4440
      Begin OcxLabelX.LabelX LblxMonto 
         Height          =   420
         Left            =   2940
         TabIndex        =   6
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Caption         =   "0.00"
         Bold            =   -1  'True
         Alignment       =   1
      End
      Begin VB.ComboBox CmbCheque 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   2325
      End
      Begin VB.ComboBox CmbBancos 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   3300
      End
      Begin OcxLabelX.LabelX LblxMoneda 
         Height          =   420
         Left            =   855
         TabIndex        =   8
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   -1  'True
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX lblxDisponible 
         Height          =   420
         Left            =   2040
         TabIndex        =   11
         Top             =   1800
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Caption         =   "0.00"
         Bold            =   -1  'True
         Alignment       =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Disponible :"
         Height          =   195
         Left            =   1200
         TabIndex        =   12
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1335
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Monto :"
         Height          =   210
         Left            =   2340
         TabIndex        =   5
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cheque :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco : "
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmBuscaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nChequeEstado As ChequeEstado
Private vnMoneda As Integer
Dim MatDatos(5) As String
Dim lnCred As Long
'By Capi 15042008
Dim lmInicio(3) As Integer
'' *** PEAC 20090323
Dim lnIncluyeDevConvenio As Integer
Dim lnSaldoCheque As Double
Dim lnTipoDoc As Integer
Dim lsNumeDoc As String
Dim lnMontoDisponible As Currency
Dim lsCodIns As String
'' ************
'**********RECO 2013-07-23**********
Public pnMontoDisponible As Double
'*************END RECO**************


Public Function BuscaCheque(ByVal pnChqEstado As ChequeEstado, Optional ByVal pnMoneda As Integer = -1, Optional ByVal pnCred As Integer = -1, Optional ByVal pnIncluyeDevConvenio As Integer = 0, Optional ByVal plsCodIns As String = "") As Variant
    lsCodIns = plsCodIns
    nChequeEstado = pnChqEstado
    vnMoneda = pnMoneda
    lnCred = pnCred
    lnIncluyeDevConvenio = pnIncluyeDevConvenio
    
    Call CargaBancos
    If CmbBancos.ListCount > 0 Then
        CmbBancos.ListIndex = 0
    End If
    Me.Show 1
    BuscaCheque = MatDatos
End Function
'MADM 20110611 - plsCodIns
Private Sub CargaCheques(ByVal psPersCod As String, ByVal pnChqEstado As ChequeEstado, ByVal pnMoneda As Integer, Optional lnCred As Long = -1, Optional pnIncluyeDevConvenio As Integer = 0)
Dim oFinan As COMDPersona.DCOMInstFinac
Dim RChq As ADODB.Recordset

    Set oFinan = New COMDPersona.DCOMInstFinac
    Set RChq = Nothing
    
    If lnCred = -1 Then
        Set RChq = oFinan.CargaChequesBanco(psPersCod, pnChqEstado, lsCodIns)
    Else
        Set RChq = oFinan.CargaChequesBancoCred(psPersCod, pnChqEstado, lsCodIns)
    End If
        
    CmbCheque.Clear
    Do While Not RChq.EOF
        If RChq!nmoneda = pnMoneda Or pnMoneda = -1 Then
            'By Capi 15042008
            'CmbCheque.AddItem Trim(RChq!cNroDoc) & Space(50) & Format(RChq!nMonto, "#0.00") & RChq!nmoneda
            
            ''*** PEAC 20090323
'            If (IIf(IsNull(RChq!nMonto), 0, RChq!nMonto) - IIf(IsNull(RChq!nMOntoUsadoCh - lnSaldoCheque), 0, RChq!nMOntoUsadoCh - lnSaldoCheque)) > 0 Then
'                CmbCheque.AddItem Trim(RChq!cNroDoc) & Space(50) & "*" & Format(RChq!nMonto, "#0.00") & "*" & Format(IIf(IsNull(RChq!nMonto), 0, RChq!nMonto) - IIf(IsNull(RChq!nMOntoUsadoCh - lnSaldoCheque), 0, RChq!nMOntoUsadoCh - lnSaldoCheque), "#0.00") & "*" & RChq!nmoneda
'            End If

'            lnSaldoCheque = IIf(pnIncluyeDevConvenio = 1, IIf(IsNull(RChq!nMonDevCon), 0, RChq!nMonDevCon), 0)
            'MADM 20110701
            lnSaldoCheque = IIf(IsNull(RChq!nMonDevCon), 0, RChq!nMonDevCon)
            'END MADM
            If (IIf(IsNull(RChq!nMonto), 0, RChq!nMonto) - IIf(IsNull(RChq!nMOntoUsadoCh), 0, RChq!nMOntoUsadoCh) - IIf(IsNull(lnSaldoCheque), 0, lnSaldoCheque)) > 0 Then
                CmbCheque.AddItem Trim(RChq!cNroDoc) & Space(50) & "*" & Format(RChq!nMonto, "#0.00") & "*" & Format(IIf(IsNull(RChq!nMonto), 0, RChq!nMonto) - IIf(IsNull(RChq!nMOntoUsadoCh), 0, RChq!nMOntoUsadoCh) - IIf(IsNull(lnSaldoCheque), 0, lnSaldoCheque), "#0.00") & "*" & RChq!nmoneda & "*" & Space(30) & IIf(IsNull(RChq!nTpoDoc), 0, RChq!nTpoDoc)
            End If
            '************* FIN PEAC
            '
        End If
        RChq.MoveNext
    Loop
   
    RChq.Close
    
    'Set RChq = oFinan.CargaChequesBanco(psPersCod, gChqEstValorizado)
    'Set oFinan = Nothing
    'Do While Not RChq.EOF
    '    If RChq!nMoneda = pnMoneda Or pnMoneda = -1 Then
    '        CmbCheque.AddItem Trim(RChq!cNroDoc) & Space(50) & Format(RChq!nMonto, "#0.00") & RChq!nMoneda
    '    End If
    '    RChq.MoveNext
    'Loop
    'RChq.Close
    
    Set RChq = Nothing
End Sub

Private Sub CargaBancos()
Dim oFinan As COMDPersona.DCOMInstFinac
Dim R As ADODB.Recordset
    
    CmbBancos.Clear
    Set oFinan = New COMDPersona.DCOMInstFinac
    Set R = oFinan.RecuperaBancos(False, lsCodIns)
    Do While Not R.EOF
        CmbBancos.AddItem PstaNombre(R!cPersNombre) & Space(50) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oFinan = Nothing
End Sub

Private Sub CmbBancos_Click()
    If CmbBancos.ListIndex <> -1 Then
         Call CargaCheques(Trim(Right(CmbBancos.Text, 30)), nChequeEstado, vnMoneda, 1, lnIncluyeDevConvenio)
    End If
End Sub

Private Sub CmbBancos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbBancos.ListIndex <> -1 Then
            Call CargaCheques(Trim(Right(CmbBancos.Text, 30)), nChequeEstado, vnMoneda)
        End If
        CmbCheque.SetFocus
    End If
End Sub

Private Sub CmbCheque_Click()
Dim sCad As String
Dim oFinan As COMDPersona.DCOMInstFinac
Set oFinan = New COMDPersona.DCOMInstFinac
    If CmbCheque.ListIndex <> -1 Then
        'By Capi 15042008
        Dim x As Integer
        For x = 0 To 2
            If x = 0 Then
                lmInicio(x) = InStr(x + 1, Trim(CmbCheque.Text), "*", vbTextCompare) + 1
            Else
                lmInicio(x) = InStr(lmInicio(x - 1) + 1, Trim(CmbCheque.Text), "*", vbTextCompare) + 1
            End If
        Next
                
                
'IIf(IsNull(RChq!nMonDevCon), 0, RChq!nMonDevCon)
                
                
        sCad = Trim(Mid(CmbCheque.Text, Len(CmbCheque.Text) - 63, 31))
        'By Capi 15042008
        'LblxMonto.Caption = Trim(Right(CmbCheque.Text, 30))
        'LblxMonto.Caption = Mid(LblxMonto.Caption, 1, Len(LblxMonto.Caption) - 1)
        LblxMonto.Caption = Mid(Trim(CmbCheque.Text), lmInicio(0), lmInicio(1) - lmInicio(0) - 1)
        
        'MADM 20110328 - valida si un cheque es repetido - banco
        If oFinan.ExisteRegistroCheque2(Trim(Right(CmbCheque.Text, 4)), Trim(Mid(CmbCheque.Text, 1, 15))) Then
            lnMontoDisponible = Mid(Trim(CmbCheque.Text), lmInicio(1), lmInicio(2) - lmInicio(1) - 1) - oFinan.CargaChequeMontoUsadoInstitucion(Trim(Right(CmbCheque.Text, 4)), Trim(Mid(CmbCheque.Text, 1, 15)), Trim(Right(CmbBancos.Text, 30)), lsCodIns)
        Else
            lnMontoDisponible = Mid(Trim(CmbCheque.Text), lmInicio(1), lmInicio(2) - lmInicio(1) - 1) - oFinan.CargaChequeMontoUsado(Trim(Right(CmbCheque.Text, 4)), Trim(Mid(CmbCheque.Text, 1, 15)))
        End If
        'END MADM
        
        'lnMontoDisponible = Mid(Trim(CmbCheque.Text), lmInicio(1), lmInicio(2) - lmInicio(1) - 1) - oFinan.CargaChequeMontoUsado(Trim(Right(CmbCheque.Text, 4)), Trim(Mid(CmbCheque.Text, 1, 15)))
        lblxDisponible.Caption = Format(lnMontoDisponible, "#0.00")
        'If Right(sCad, 1) = "1" Then
        
        Set oFinan = Nothing
        'lnTipoDoc
        'lsNumeDoc
        If Mid(Trim(CmbCheque.Text), lmInicio(2), 1) = "1" Then
            LblxMoneda.Caption = "SOLES"
            LblxMonto.Resalte = eNegro
        Else
            LblxMoneda.Caption = "DOLARES"
            LblxMonto.Resalte = eVerde
        End If
    End If
End Sub

Private Sub CmbCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    
    If Me.CmbCheque.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Cheque", vbInformation, "Aviso"
        CmbCheque.SetFocus
        Exit Sub
    End If
    'By Capi 15042008
    'MatDatos(0) = Mid(CmbBancos.Text, 1, 45) 'Descripcion Banco
    'MatDatos(0) = Mid(Trim(CmbCheque.Text), lmInicio(1), lmInicio(2) - lmInicio(1) - 1) ' Monto Disponible
    MatDatos(0) = lnMontoDisponible
    MatDatos(1) = Trim(Right(CmbBancos.Text, 30)) 'Codigo de Persona
    'MatDatos(2) = Right(CmbCheque.Text, 1) 'Moneda
    MatDatos(2) = Mid(Trim(CmbCheque.Text), lmInicio(2), 1) 'Moneda
    'MatDatos(3) = Mid(Trim(Right(CmbCheque.Text, 30)), 1, Len(Trim(Right(CmbCheque.Text, 30))) - 1) 'Monto
    MatDatos(3) = lnMontoDisponible
    'MatDatos(3) = Mid(Trim(CmbCheque.Text), lmInicio(0), lmInicio(1) - lmInicio(0) - 1)
    MatDatos(4) = Trim(Mid(CmbCheque.Text, 1, 30)) 'Nro de Cheque
    '******************RECO 2013-07-23******************
    pnMontoDisponible = lnMontoDisponible
    '****************END RECO***************************

    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    MatDatos(0) = ""
    MatDatos(1) = ""
    MatDatos(2) = ""
    MatDatos(3) = ""
    MatDatos(4) = ""
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaBancos
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
