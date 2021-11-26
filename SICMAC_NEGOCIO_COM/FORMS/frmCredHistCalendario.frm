VERSION 5.00
Begin VB.Form frmCredHistCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de Calendarios"
   ClientHeight    =   6015
   ClientLeft      =   1245
   ClientTop       =   3660
   ClientWidth     =   11280
   Icon            =   "frmCredHistCalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11280
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   90
      TabIndex        =   2
      Top             =   5235
      Width           =   11130
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nueva Busqueda"
         Height          =   405
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   1530
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   9645
         TabIndex        =   3
         Top             =   210
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   90
      TabIndex        =   1
      Top             =   15
      Width           =   11070
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   480
         Left            =   3885
         TabIndex        =   12
         Top             =   555
         Width           =   1515
         Begin VB.CheckBox ChkMiViv 
            Alignment       =   1  'Right Justify
            Caption         =   "Mi Vivienda"
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.ComboBox CboCuota 
         Height          =   315
         Left            =   8670
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   720
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3675
         _extentx        =   6482
         _extenty        =   873
         texto           =   "Credito :"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Dias Atraso: "
         Height          =   240
         Left            =   9480
         TabIndex        =   15
         Top             =   750
         Width           =   855
      End
      Begin VB.Label LblDiasAtraso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   10440
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Nro Calendario :"
         Height          =   240
         Left            =   7470
         TabIndex        =   10
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label LblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   6315
         TabIndex        =   9
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo :"
         Height          =   240
         Left            =   5505
         TabIndex        =   8
         Top             =   720
         Width           =   750
      End
      Begin VB.Label LblTitu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4560
         TabIndex        =   7
         Top             =   270
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Titular :"
         Height          =   240
         Left            =   3855
         TabIndex        =   6
         Top             =   270
         Width           =   570
      End
   End
   Begin SICMACT.FlexEdit FECalend 
      Height          =   3825
      Left            =   90
      TabIndex        =   0
      Top             =   1410
      Width           =   11070
      _extentx        =   19526
      _extenty        =   6747
      cols0           =   16
      fixedcols       =   0
      highlight       =   1
      allowuserresizing=   1
      encabezadosnombres=   $"frmCredHistCalendario.frx":030A
      encabezadosanchos=   "400-1200-1200-1200-700-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
      font            =   "frmCredHistCalendario.frx":0391
      font            =   "frmCredHistCalendario.frx":03BD
      font            =   "frmCredHistCalendario.frx":03E9
      font            =   "frmCredHistCalendario.frx":0415
      font            =   "frmCredHistCalendario.frx":0441
      fontfixed       =   "frmCredHistCalendario.frx":046D
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   1
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "No"
      colwidth0       =   405
      rowheight0      =   300
      forecolor       =   -2147483630
      forecolorfixed  =   -2147483635
      cellforecolor   =   -2147483630
   End
End
Attribute VB_Name = "frmCredHistCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bPagoCuotas As Boolean
Private sCtaCod As String

Public Sub PagoCuotas(ByVal psCtaCod As String)
    bPagoCuotas = True
    ActxCta.NroCuenta = psCtaCod
    
    Call ActxCta_KeyPress(13)
    
    CboCuota.Enabled = False
    CmdNuevo.Enabled = False
    Me.Show 1
End Sub

Public Sub Inicio()
    bPagoCuotas = False
    Me.Show 1
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)

'Dim oCred As COMDCredito.DCOMCalendario
Dim oCredDat As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset

Dim dFecPago As String
Dim I As Integer
Dim Pos As Integer

Dim rsCredVig As ADODB.Recordset
Dim rsCalend As ADODB.Recordset

    If KeyAscii = 13 Then
        
        Set oCredDat = New COMDCredito.DCOMCredito
        'Set R = oCredDat.RecuperaDatosCreditoVigente(ActxCta.NroCuenta)
        Call oCredDat.CargarHistoriaCredito(ActxCta.NroCuenta, bPagoCuotas, rsCredVig, rsCalend)
        Set oCredDat = Nothing
        
        If rsCredVig.RecordCount > 0 Then
            
            LblTitu.Caption = Space(3) & PstaNombre(rsCredVig!cPersNombre)
            ChkMiViv.value = rsCredVig!bMiVivienda
            LblMonto.Caption = Format(rsCredVig!nMontoCol, "#0.00")
            
            CboCuota.Clear
            For I = 1 To rsCredVig!nNroCalen
                CboCuota.AddItem I
            Next I
            If Not bPagoCuotas Then
               CboCuota.ListIndex = 0
               LblDiasAtraso.Visible = False
               Label4.Visible = False
            Else
               Pos = I - 1
               CboCuota.ListIndex = IndiceListaCombo(CboCuota, Pos)
               LblDiasAtraso.Visible = True
               Label4.Visible = True
               LblDiasAtraso.Visible = Format(rsCredVig!nDiasAtraso, "0")
            End If
            'Set oCred = New COMDCredito.DCOMCalendario
            'If bPagoCuotas = False Then
            '    Set R = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text))
            'Else
            '    Set R = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text), , , , True)
            'End If
            'Set oCred = Nothing
            LimpiaFlex FECalend
            Do While Not rsCalend.EOF
                If rsCalend.Bookmark <> 1 Then
                    FECalend.AdicionaFila
                End If
                If rsCalend!nColocCalendEstado = 1 Then
                    'FECalend.ForeColorRow vbRed
                    FECalend.BackColorRow vbYellow
                End If
                
                'Agregado por LMMD
                If rsCalend!nColocCalendEstado = 0 And DateDiff("d", rsCalend!dvenc, gdFecSis) >= 0 Then
                    FECalend.BackColorRow vbRed
                    FECalend.ForeColorRow vbWhite
                End If
                
                If IsNull(rsCalend!dPago) Then
                    dFecPago = ""
                Else
                    If Year(rsCalend!dPago) = "1900" Then
                        dFecPago = ""
                    Else
                        dFecPago = Format(rsCalend!dPago, "dd/mm/yyyy")
                    End If
                End If
                
                FECalend.TextMatrix(rsCalend.Bookmark, 0) = Trim(Str(rsCalend!nCuota))
                FECalend.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dvenc, "dd/mm/yyyy")
                FECalend.TextMatrix(rsCalend.Bookmark, 2) = dFecPago
                FECalend.TextMatrix(rsCalend.Bookmark, 3) = Format(rsCalend!nCapital + rsCalend!nIntComp + rsCalend!nIntGracia + rsCalend!nIntMor + rsCalend!nIntReprog + rsCalend!nIntSuspenso + rsCalend!nGasto + rsCalend!nIntCompVenc, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 4) = IIf(rsCalend!nColocCalendEstado = 0, "Pend.", "Cancel.")
                FECalend.TextMatrix(rsCalend.Bookmark, 5) = Format(rsCalend!nCapital, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 6) = Format(rsCalend!nCapitalPag, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 7) = Format(rsCalend!nIntComp + rsCalend!nIntGracia + rsCalend!nIntReprog + rsCalend!nIntSuspenso, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 8) = Format(rsCalend!nIntCompPag + rsCalend!nIntGraciaPag + rsCalend!nIntReprogPag + rsCalend!nIntSuspensoPag, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 9) = Format(rsCalend!nIntMor, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 10) = Format(rsCalend!nIntMorPag, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 11) = Format(rsCalend!nIntCompVenc, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 12) = Format(rsCalend!nIntCompVencPag, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 13) = Format(rsCalend!nGasto, "#0.00")
                FECalend.TextMatrix(rsCalend.Bookmark, 14) = Format(rsCalend!nGastoPag, "#0.00")
                If bPagoCuotas = True Then
                    FECalend.TextMatrix(rsCalend.Bookmark, 15) = IIf(IsNull(rsCalend!Cuser), "", rsCalend!Cuser)
                End If
                rsCalend.MoveNext
            Loop
            ActxCta.Enabled = False
            FECalend.Enabled = True
        Else
            MsgBox "No Se pudo encontrar el Credito o no esta Vigente"
            'R.Close
            Exit Sub
        End If
    End If
End Sub

Private Sub CboCuota_Click()
Dim oCred As COMDCredito.DCOMCalendario
Dim R As ADODB.Recordset

    Set oCred = New COMDCredito.DCOMCalendario
    Set R = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text), , True)
    Set oCred = Nothing
    LimpiaFlex FECalend
    Do While Not R.EOF
        If R.Bookmark <> 1 Then
            FECalend.AdicionaFila
        End If
        FECalend.TextMatrix(R.Bookmark, 0) = Trim(Str(R!nCuota))
        FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dvenc, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 2) = Format(R!dPago, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 3) = Format(R!nCapital + R!nIntComp + R!nIntGracia + R!nIntMor + R!nIntReprog + R!nIntSuspenso + R!nGasto + R!nIntCompVenc, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 4) = IIf(R!nColocCalendEstado = 0, "Pend.", "Cancel.")
        FECalend.TextMatrix(R.Bookmark, 5) = Format(R!nCapital, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 6) = Format(R!nCapitalPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 7) = Format(R!nIntComp + R!nIntGracia + R!nIntReprog + R!nIntSuspenso, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 8) = Format(R!nIntCompPag + R!nIntGraciaPag + R!nIntReprogPag + R!nIntSuspensoPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 9) = Format(R!nIntMor, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 10) = Format(R!nIntMorPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 11) = Format(R!nIntCompVenc, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 12) = Format(R!nIntCompVencPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 13) = Format(R!nGasto, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 14) = Format(R!nGastoPag, "#0.00")
        R.MoveNext
    Loop
    R.Close
End Sub

Private Sub cmdNuevo_Click()
    ActxCta.Enabled = True
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LimpiaFlex FECalend
    LblMonto.Caption = "0.00"
    LblTitu.Caption = ""
    CboCuota.Clear
    ChkMiViv.value = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Not bPagoCuotas Then
       ActxCta.NroCuenta = ""
       ActxCta.CMAC = gsCodCMAC
       ActxCta.Age = gsCodAge
    End If
    CentraForm Me
End Sub
