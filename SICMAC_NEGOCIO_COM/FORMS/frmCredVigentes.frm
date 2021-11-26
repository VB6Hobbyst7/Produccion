VERSION 5.00
Begin VB.Form frmCredVigentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creditos Vigentes"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   Icon            =   "frmCredVigentes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   2115
      TabIndex        =   1
      Top             =   2655
      Width           =   1605
   End
   Begin SICMACT.FlexEdit FECreditosVig 
      Height          =   1665
      Left            =   45
      TabIndex        =   0
      Top             =   885
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   2937
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "--Credito-Atraso-Monto"
      EncabezadosAnchos=   "300-350-2000-1200-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-4-0-0-0"
      BackColor       =   14286847
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-R-R-R"
      FormatosEdit    =   "0-0-3-3-2"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
      CellBackColor   =   14286847
   End
   Begin VB.Label LblCodCred 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "111022010000002356"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   780
      TabIndex        =   5
      Top             =   90
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   "Credito :"
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblNomCli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SILVA ESTRADA NAPOLEON "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1590
      TabIndex        =   3
      Top             =   450
      Width           =   3915
   End
   Begin VB.Label LblCodCli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   75
      TabIndex        =   2
      Top             =   450
      Width           =   1485
   End
End
Attribute VB_Name = "frmCredVigentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MatCredVig() As String
Public Function Inicio(ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psCtaCod As String, ByVal pMatCredVig As Variant) As Variant

    LblCodCli.Caption = " " & psPersCod
    LblNomCli.Caption = "  " & psPersNombre
    LblCodCred.Caption = " " & psCtaCod
    
    If IsArray(pMatCredVig) Then
        If UBound(pMatCredVig) > 0 Then
            MatCredVig = pMatCredVig
        Else
            ReDim MatCredVig(0)
        End If
    Else
        ReDim MatCredVig(0)
    End If
    'FRHU 20140422 ERS015-2014
    Call CargaDatosNew
    'Call CargaDatos
    'FRHU 20140422 ERS015-2014
    Me.Show 1
    
    Inicio = MatCredVig
    
End Function
'FRHU 20140422 ERS015-2014
Private Sub CargaDatosNew()
    Dim oNegCred As New COMNCredito.NCOMCredito
    Dim rsVigentes As ADODB.Recordset
    Set rsVigentes = oNegCred.CargarCreditosVigentesNew(LblCodCli.Caption, LblCodCred.Caption, Mid(Trim(LblCodCred.Caption), 9, 1))
    LimpiaFlex FECreditosVig
    Do While Not rsVigentes.EOF
        FECreditosVig.AdicionaFila , , True
        If CStr(rsVigentes!Ampliado) <> "" Then
            FECreditosVig.TextMatrix(rsVigentes.Bookmark, 1) = "1"
        End If
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 2) = rsVigentes!cCtaCod
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 3) = rsVigentes!nDiasAtraso
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 4) = Format(rsVigentes!nSaldo + oNegCred.InteresGastosAFecha(rsVigentes!cCtaCod, gdFecSis), "#0.00")
        rsVigentes.MoveNext
    Loop
    Set oNegCred = Nothing
End Sub
'FIN FRHU 20140422 ERS015-2014
Private Sub CargaDatos()
'Dim R As ADODB.Recordset
Dim oNegCred As COMNCredito.NCOMCredito
'Dim oCredito As COMDCredito.DCOMCredito
Dim rsVigentes As ADODB.Recordset
Dim rsGrabados As ADODB.Recordset
Dim i, K As Integer

    LimpiaFlex FECreditosVig
    'Set oCredito = New COMDCredito.DCOMCredito
    Set oNegCred = New COMNCredito.NCOMCredito
    Call oNegCred.CargarCreditosVigentes(LblCodCred.Caption, LblCodCli.Caption, Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc), Mid(LblCodCred.Caption, 9, 1), rsVigentes, rsGrabados)
    'Set R = oCredito.RecuperaCreditosVigentes(LblCodCli.Caption, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc), Mid(LblCodCred.Caption, 9, 1))
    Do While Not rsVigentes.EOF
        FECreditosVig.AdicionaFila , , True
        'FECreditosVig.TextMatrix(r.Bookmark, 0) = r.Bookmark
        'FECreditosVig.TextMatrix(rsVigentes.Bookmark, 1) = "1"
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 2) = rsVigentes!cCtaCod
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 3) = rsVigentes!nDiasAtraso
        FECreditosVig.TextMatrix(rsVigentes.Bookmark, 4) = Format(rsVigentes!nSaldo + oNegCred.InteresGastosAFecha(rsVigentes!cCtaCod, gdFecSis), "#0.00")
        rsVigentes.MoveNext
    Loop
    Set oNegCred = Nothing
    'FECreditosVig.Rows = r.RecordCount + 1
    'R.Close
    'Set R = Nothing
    
    'Set R = oCredito.RecuperaCreditosVigentesGrabados(LblCodCred.Caption)
    If UBound(MatCredVig) = 0 Then
        If rsGrabados.RecordCount > 0 Then
        ReDim MatCredVig(rsGrabados.RecordCount, 2)
    
        
            If Trim(MatCredVig(0, 0)) = "" Then
                Do While Not rsGrabados.EOF
                        MatCredVig(rsGrabados.Bookmark - 1, 0) = Me.LblCodCred.Caption
                        MatCredVig(rsGrabados.Bookmark - 1, 1) = Trim(rsGrabados!cCtaCodRef)
                    rsGrabados.MoveNext
                Loop
            End If
        End If
    End If
    'R.Close
    'Set oNegCred = Nothing
    'Set oCredito = Nothing
    
    If IsArray(MatCredVig) Then
        For K = 0 To UBound(MatCredVig) - 1
            For i = 1 To FECreditosVig.Rows - 1
                If Trim(FECreditosVig.TextMatrix(i, 2)) = Trim(MatCredVig(K, 1)) Then
                    FECreditosVig.TextMatrix(i, 1) = "1"
                End If
            Next i
        Next K
    End If
    
End Sub

Private Sub CmdAceptar_Click()
Dim i As Integer
Dim nCont As Integer

    nCont = 0
    ReDim MatCredVig(FECreditosVig.Rows - 1, 2)
    
    For i = 1 To FECreditosVig.Rows - 1
        If FECreditosVig.TextMatrix(i, 1) = "." Then
            nCont = nCont + 1
            MatCredVig(i - 1, 0) = Trim(LblCodCred.Caption)
            MatCredVig(i - 1, 1) = Trim(FECreditosVig.TextMatrix(i, 2))
        End If
    Next i
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
