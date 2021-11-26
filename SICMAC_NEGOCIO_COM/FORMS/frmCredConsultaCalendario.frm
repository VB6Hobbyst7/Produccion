VERSION 5.00
Begin VB.Form frmCredHistCalendario 
   Caption         =   "Historial de Calendarios"
   ClientHeight    =   6000
   ClientLeft      =   360
   ClientTop       =   1830
   ClientWidth     =   11280
   Icon            =   "frmCredHistCalendario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6000
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
         Left            =   4365
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
         Left            =   9150
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
         _ExtentX        =   6482
         _ExtentY        =   873
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Nro Calendario :"
         Height          =   240
         Left            =   7950
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
         Left            =   6795
         TabIndex        =   9
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo :"
         Height          =   240
         Left            =   5985
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
         Left            =   5040
         TabIndex        =   7
         Top             =   270
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Titular :"
         Height          =   240
         Left            =   4335
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
      _ExtentX        =   19526
      _ExtentY        =   6747
      Cols0           =   14
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "No-Fecha-Cuota-Estado-Capital-Capital Pag-Interes-Int. Pag-Mora-Mora Pag-Int Venc-Int Venc Pag-Gastos-Gasto Pag"
      EncabezadosAnchos=   "400-1200-1200-700-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "No"
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColor       =   -2147483630
      ForeColorFixed  =   -2147483635
      CellForeColor   =   -2147483630
   End
End
Attribute VB_Name = "frmCredHistCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
Dim oCred As Dcalendario
Dim oCredDat As DCredito
Dim R As ADODB.Recordset
Dim i As Integer

    If KeyAscii = 13 Then
        
        Set oCredDat = New DCredito
        Set R = oCredDat.RecuperaDatosCreditoVigente(ActxCta.NroCuenta)
        Set oCredDat = Nothing
        
        If R.RecordCount > 0 Then
            
            LblTitu.Caption = Space(3) & PstaNombre(R!cPersNombre)
            ChkMiViv.value = R!bMiVivienda
            LblMonto.Caption = Format(R!nMontoCol, "#0.00")
            
            CboCuota.Clear
            For i = 1 To R!nNroCalen
                CboCuota.AddItem i
            Next i
            CboCuota.ListIndex = 0
            Set oCred = New Dcalendario
            Set R = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text))
            Set oCred = Nothing
            LimpiaFlex FECalend
            Do While Not R.EOF
                If R.Bookmark <> 1 Then
                    FECalend.AdicionaFila
                End If
                FECalend.TextMatrix(R.Bookmark, 0) = Trim(Str(R!nCuota))
                FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dVenc, "dd/mm/yyyy")
                FECalend.TextMatrix(R.Bookmark, 2) = Format(R!nCapital + R!nIntComp + R!nIntGracia + R!nIntMor + R!nIntReprog + R!nIntSuspenso + R!nGasto + R!nIntCompVenc, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 3) = IIf(R!nColocCalendEstado = 0, "Pend.", "Cancel.")
                FECalend.TextMatrix(R.Bookmark, 4) = Format(R!nCapital, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 5) = Format(R!nCapitalPag, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 6) = Format(R!nIntComp + R!nIntGracia + R!nIntReprog + R!nIntSuspenso, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 7) = Format(R!nIntCompPag + R!nIntGraciaPag + R!nIntReprogPag + R!nIntSuspensoPag, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 8) = Format(R!nIntMor, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 9) = Format(R!nIntMorPag, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 10) = Format(R!nIntCompVenc, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 11) = Format(R!nIntCompVencPag, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 12) = Format(R!nGasto, "#0.00")
                FECalend.TextMatrix(R.Bookmark, 13) = Format(R!nGastoPag, "#0.00")
                R.MoveNext
            Loop
            ActxCta.Enabled = False
            FECalend.Enabled = True
        Else
            MsgBox "No Se pudo encontrar el Credito o no esta Vigente"
            R.Close
            Exit Sub
        End If
    End If
End Sub

Private Sub CboCuota_Click()
Dim oCred As Dcalendario
Dim R As ADODB.Recordset

    Set oCred = New Dcalendario
    Set R = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text), , True)
    Set oCred = Nothing
    LimpiaFlex FECalend
    Do While Not R.EOF
        If R.Bookmark <> 1 Then
            FECalend.AdicionaFila
        End If
        FECalend.TextMatrix(R.Bookmark, 0) = Trim(Str(R!nCuota))
        FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dVenc, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 2) = Format(R!nCapital + R!nIntComp + R!nIntGracia + R!nIntMor + R!nIntReprog + R!nIntSuspenso + R!nGasto + R!nIntCompVenc, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 3) = IIf(R!nColocCalendEstado = 0, "Pend.", "Cancel.")
        FECalend.TextMatrix(R.Bookmark, 4) = Format(R!nCapital, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 5) = Format(R!nCapitalPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 6) = Format(R!nIntComp + R!nIntGracia + R!nIntReprog + R!nIntSuspenso, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 7) = Format(R!nIntCompPag + R!nIntGraciaPag + R!nIntReprogPag + R!nIntSuspensoPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 8) = Format(R!nIntMor, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 9) = Format(R!nIntMorPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 10) = Format(R!nIntCompVenc, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 11) = Format(R!nIntCompVencPag, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 12) = Format(R!nGasto, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 13) = Format(R!nGastoPag, "#0.00")
        R.MoveNext
    Loop
    R.Close
End Sub

Private Sub CmdNuevo_Click()
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

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    CentraForm Me
End Sub
