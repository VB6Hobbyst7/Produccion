VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImprimirCupones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Cupones de Campaña"
   ClientHeight    =   4380
   ClientLeft      =   2730
   ClientTop       =   3765
   ClientWidth     =   8865
   Icon            =   "frmImprimirCupones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8865
   Begin VB.CommandButton cmdNext 
      Height          =   330
      Left            =   8130
      Picture         =   "frmImprimirCupones.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   450
      Width           =   405
   End
   Begin VB.CommandButton cndAnt 
      Height          =   330
      Left            =   8130
      Picture         =   "frmImprimirCupones.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   75
      Width           =   405
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   7620
      TabIndex        =   5
      Top             =   3840
      Width           =   1080
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   105
      TabIndex        =   4
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   2910
      Left            =   105
      TabIndex        =   3
      Top             =   840
      Width           =   8610
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   105
         TabIndex        =   27
         Top             =   210
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   3096
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-TITULAR-DNI"
         EncabezadosAnchos=   "600-5500-1800"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   2415
         TabIndex        =   25
         Top             =   3015
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtHasta 
         Height          =   285
         Left            =   4035
         TabIndex        =   22
         Top             =   3225
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtDesde 
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   3225
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox chkCuenta 
         Caption         =   "Por Cuenta"
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
         Left            =   195
         TabIndex        =   20
         Top             =   2970
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox chkCupon 
         Caption         =   "Un Cupón"
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
         Left            =   195
         TabIndex        =   17
         Top             =   2940
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox chkRango 
         Caption         =   "De un Rango"
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
         Left            =   195
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresión"
         Height          =   795
         Left            =   6255
         TabIndex        =   8
         Top             =   2010
         Visible         =   0   'False
         Width           =   2220
         Begin VB.OptionButton optImpresion 
            Caption         =   "Pantalla"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   11
            Top             =   285
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Impresora"
            Height          =   270
            Index           =   1
            Left            =   495
            TabIndex        =   10
            Top             =   480
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Archivo"
            Height          =   270
            Index           =   2
            Left            =   525
            TabIndex        =   9
            Top             =   975
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cupones:"
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
         Left            =   3465
         TabIndex        =   31
         Top             =   2565
         Width           =   1170
      End
      Begin VB.Label lblNroCupones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4935
         TabIndex        =   30
         Top             =   2475
         Width           =   1185
      End
      Begin VB.Label lblNumImp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4935
         TabIndex        =   29
         Top             =   2085
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Impresiones:"
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
         Left            =   3450
         TabIndex        =   28
         Top             =   2175
         Width           =   1440
      End
      Begin VB.Label lbletiq3 
         AutoSize        =   -1  'True
         Caption         =   "Nro:"
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
         Left            =   1965
         TabIndex        =   26
         Top             =   2955
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblEtiq2 
         AutoSize        =   -1  'True
         Caption         =   "al:"
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
         Left            =   3780
         TabIndex        =   24
         Top             =   3270
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblEtiq1 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
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
         Left            =   1950
         TabIndex        =   23
         Top             =   3270
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label LblRangoFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1395
         TabIndex        =   15
         Top             =   2490
         Width           =   1890
      End
      Begin VB.Label lblRangoIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1395
         TabIndex        =   14
         Top             =   2055
         Width           =   1890
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rango Final:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   2565
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rango Inicial:"
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
         Left            =   90
         TabIndex        =   12
         Top             =   2160
         Width           =   1200
      End
   End
   Begin VB.Frame fraCuenta 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4230
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   225
         Width           =   375
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "108"
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   165
      Top             =   2355
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNumSorteo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5145
      TabIndex        =   7
      Top             =   135
      Width           =   2910
   End
   Begin VB.Label Label2 
      Caption         =   "Nro Sorteo:"
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
      Left            =   4410
      TabIndex        =   6
      Top             =   195
      Width           =   645
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmImprimirCupones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1
Dim MatSorteos()  As String
Dim i As Integer
Dim NumSort As Integer
Dim NumProd As Producto
Public gsAlcance As String



Public Sub CargaMatriz(ByVal cAlcance As String)
Dim oSorteo As Dsorteo, rsTemp As Recordset
    
    Set oSorteo = New Dsorteo
    Set rsTemp = oSorteo.GetSorteos(Year(gdFecSis), gsAlcance)
    
    If Not rsTemp.EOF Then
          ReDim MatSorteos(rsTemp.RecordCount)
          i = 1
          While Not rsTemp.EOF
                  MatSorteos(i) = rsTemp!cnumsorteo
                      i = i + 1
                  rsTemp.MoveNext
                  If rsTemp.EOF Then i = i - 1
          Wend
          NumSort = UBound(MatSorteos)
    End If
    
End Sub
Private Sub Limpiar()
    
    'lblTitular.Caption = ""
    lblRangoIni.Caption = ""
    lblRangoFin.Caption = ""
    txtDesde.Text = ""
    txtHasta.Text = ""
    txtNum.Text = ""
    chkCuenta.value = vbChecked
End Sub

Private Sub chkCuenta_Click()
  If chkCuenta.value = vbUnchecked Then
            'chkCupon.value = vbChecked
            'fraCupon.Visible = False
    Else
        If chkCupon.value = vbChecked Then
            chkCupon.value = vbUnchecked
        ElseIf chkRango.value = vbChecked Then
            chkRango.value = vbUnchecked
            
        End If
            'chkCupon.value = vbUnchecked
            'FraRango.Visible = False
           
            lbletiq3.Visible = False
            txtNum.Visible = False
            
            lblEtiq1.Visible = False
            lblEtiq2.Visible = False
            txtHasta.Visible = False
            txtDesde.Visible = False
            
    End If
End Sub

Private Sub chkCupon_Click()
   If chkCupon.value = vbUnchecked Then
            'chkCupon.value = vbChecked
            'fraCupon.Visible = False
            lbletiq3.Visible = False
            txtNum.Visible = False
    Else
        If chkRango.value = vbChecked Then
            chkRango.value = vbUnchecked
        ElseIf chkCuenta.value = vbChecked Then
            chkCuenta.value = vbUnchecked
        End If
            'chkCupon.value = vbUnchecked
            'FraRango.Visible = False
            'fraCupon.Visible = True
            
            lbletiq3.Visible = True
            txtNum.Visible = True
            
            lblEtiq1.Visible = False
            lblEtiq2.Visible = False
            txtHasta.Visible = False
            txtDesde.Visible = False
    End If
End Sub

Private Sub chkRango_Click()
    If chkRango.value = vbUnchecked Then
            ' chkRango.value = vbChecked
           '* FraRango.Visible = False
            lblEtiq1.Visible = False
            lblEtiq2.Visible = False
            txtHasta.Visible = False
            txtDesde.Visible = False
            
    Else
            If chkCupon.value = vbChecked Then
                chkCupon.value = vbUnchecked
            ElseIf chkCuenta.value = vbChecked Then
                chkCuenta.value = vbUnchecked
            End If
            
           ' chkRango.value = vbUnchecked
           ' fraCupon.Visible = False
            '*FraRango.Visible = True
            lbletiq3.Visible = False
            txtNum.Visible = False
            
            lblEtiq1.Visible = True
            lblEtiq2.Visible = True
            txtHasta.Visible = True
            txtDesde.Visible = True
            
            
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As UPersona
txtCuenta.Age = "": txtCuenta.Cuenta = ""
Limpiar
Set clsPers = New UPersona

Set clsPers = frmBuscaPersona.Inicio


If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As Recordset
    Dim clsCap As NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuentas
    sPers = clsPers.sPerscod
    Set clsCap = New NCapMantenimiento
    
        Set rsPers = clsCap.GetCuentasPersona(sPers, NumProd, , , , , "")
    If rsPers.EOF Or rsPers.BOF Then
        Set rsPers = clsCap.GetCuentasPersona(sPers, NumProd, , , , , "")
    End If
    
    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = frmCapMantenimientoCtas.Inicia
        If clsCuenta.sCtaCod <> "" Then
            txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
            txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
            txtCuenta.SetFocusCuenta
            SendKeys "{Enter}"
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdImprimir_Click()
Dim lnAge As Integer, sDato1 As String, sDato2 As String, sDato3 As String, sDato4 As String, sDato5 As String

Dim oPrevio As Previo.clsPrevio, lscadena As String
Dim orep As nCaptaReportes, oSorteo As Dsorteo
Set orep = New nCaptaReportes
Set oPrevio = New Previo.clsPrevio
Set oSorteo = New Dsorteo
On Error GoTo ErrMsg
If CLng(lblNumImp.Caption) >= 1 Then
   If MsgBox("Esta seguro de reimprimir cupones???", vbYesNo + vbQuestion, "Aviso") = vbNo Then
    Exit Sub
   End If
End If
    
    sDato1 = Trim(Left(grdCliente.TextMatrix(1, 1), 30)) & Space(31 - Len(Trim(Left(grdCliente.TextMatrix(1, 1), 30)))) & grdCliente.TextMatrix(1, 2)
    
    If grdCliente.Rows > 2 Then
        sDato2 = Trim(Left(grdCliente.TextMatrix(2, 1), 30)) & Space(31 - Len(Trim(Left(grdCliente.TextMatrix(2, 1), 30)))) & grdCliente.TextMatrix(2, 2)
    End If
    If grdCliente.Rows > 3 Then
        sDato3 = Trim(Left(grdCliente.TextMatrix(3, 1), 30)) & Space(31 - Len(Trim(Left(grdCliente.TextMatrix(3, 1), 30)))) & grdCliente.TextMatrix(3, 2)
    End If
    If grdCliente.Rows > 4 Then
        sDato4 = Trim(Left(grdCliente.TextMatrix(4, 1), 30)) & Space(31 - Len(Trim(Left(grdCliente.TextMatrix(4, 1), 30)))) & grdCliente.TextMatrix(4, 2)
    End If
    If grdCliente.Rows > 5 Then
        sDato5 = Trim(Left(grdCliente.TextMatrix(5, 1), 30)) & Space(31 - Len(Trim(Left(grdCliente.TextMatrix(5, 1), 30)))) & grdCliente.TextMatrix(5, 2)
    End If
    
    If chkCuenta.value = vbChecked Then
        lscadena = orep.CUPONIMPRIMIR(Trim(lblNumSorteo.Caption), sDato1, lblRangoIni.Caption, lblRangoFin.Caption, txtCuenta.NroCuenta, 1, , , , sDato2, sDato3, sDato4, sDato5, Val(Me.lblNroCupones.Caption))
        
    ElseIf chkRango.value = vbChecked Then
         'lscadena = orep.CUPONIMPRIMIR(Trim(lblNumSorteo.Caption), lblTitular.Caption, lblRangoIni.Caption, LblRangoFin.Caption, txtCuenta.NroCuenta, 2, , txtDesde.Text, txtHasta.Text)
    Else
         'lscadena = orep.CUPONIMPRIMIR(Trim(lblNumSorteo.Caption), lblTitular.Caption, lblRangoIni.Caption, LblRangoFin.Caption, txtCuenta.NroCuenta, 3, txtNum.Text)
    End If
        
    If Me.optImpresion(0).value = True Then
        Set oPrevio = New Previo.clsPrevio
            oPrevio.Show lscadena, "CUPONES PARA SORTEO " & Trim(Me.lblNumSorteo.Caption), True
        Set oPrevio = Nothing
        
    Else
    
        Set oPrevio = New Previo.clsPrevio
        oPrevio.PrintSpool sLpt, lscadena, True
        Set oPrevio = Nothing
    
'        dlgGrabar.CancelError = True
'        dlgGrabar.InitDir = App.path
'        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'        dlgGrabar.ShowSave
'        If dlgGrabar.FileName <> "" Then
'           Open dlgGrabar.FileName For Output As #1
'            Print #1, lscadena
'            Close #1
'        End If
    End If
    
    Call oSorteo.ActualizaCuentaImpresion(Trim(Me.lblNumSorteo.Caption), txtCuenta.NroCuenta)
    lblNumImp.Caption = CLng(lblNumImp.Caption) + 1
    
    Set oPrevio = Nothing

    
    
   Exit Sub
ErrMsg:
     MsgBox "Ocurrió el error:" & Err.Number & " " & Err.Description, vbOKOnly, "Aviso"
     
End Sub

Private Sub cmdNext_Click()
If i < UBound(MatSorteos) And i >= LBound(MatSorteos) Then
    i = i + 1
    lblNumSorteo.Caption = MatSorteos(i)
End If


End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub cndAnt_Click()


If i > LBound(MatSorteos) Then
    i = i - 1
    lblNumSorteo.Caption = MatSorteos(i)
    
End If


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vsprod As Producto
 vsprod = NumProd
 
 If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
    Limpiar
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(vsprod, False)
        If Val(Mid(sCuenta, 6, 3)) <> vsprod Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If

End Sub

Private Sub Form_Load()

    gsAlcance = frmPreparaSorteo.gsAlcance
    CargaMatriz (gsAlcance)
    If i > 0 Then
       lblNumSorteo.Caption = MatSorteos(i)
    End If
    Me.Left = 2685
    Me.Top = 3435
End Sub



Private Sub Label9_Click()

End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
Dim oSorteo  As Dsorteo
Dim rsTemp As Recordset, rsRel As Recordset
Dim clsMant As NCapMantenimiento, i As Integer



i = 1
If KeyAscii = 13 Then
    Limpiar
    Set oSorteo = New Dsorteo
    Set rsTemp = oSorteo.GetDatosCtaSorteo(Left(Trim(lblNumSorteo.Caption), 6), txtCuenta.NroCuenta)
    If Not rsTemp.EOF Then
       'lblTitular.Caption = RSTEMP!cPersNombre
       Set clsMant = New NCapMantenimiento
      Set rsRel = clsMant.GetPersonaCuenta(txtCuenta.NroCuenta)
      grdCliente.Clear
      grdCliente.FormaCabecera
      
        Do While Not rsRel.EOF
            grdCliente.AdicionaFila
            grdCliente.TextMatrix(i, 0) = i
            grdCliente.TextMatrix(i, 1) = UCase(PstaNombre(rsRel("Nombre")))
            grdCliente.TextMatrix(i, 2) = rsRel("ID N°")
            i = i + 1
        
         rsRel.MoveNext
        Loop
       
       Me.lblNumImp.Caption = rsTemp!Nimpresiones
       Me.lblNroCupones.Caption = rsTemp!nnumtickets
       lblRangoIni.Caption = rsTemp!nRangoIni
       lblRangoFin.Caption = rsTemp!nRangoFin
    Else
     MsgBox "Esta cuenta no entra al sorteo " & Trim(lblNumSorteo.Caption), vbOKOnly + vbInformation, "Aviso"
    End If
End If

End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
 If KeyAscii <> 13 Then
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
      End If
 Else
          If txtDesde.Text <> "" Then
            If Not (CLng(txtDesde.Text) >= CLng(lblRangoIni.Caption) And CLng(txtDesde.Text) < CLng(lblRangoFin.Caption)) Then
                txtDesde.SetFocus
                Exit Sub
            End If
          End If
        txtHasta.SetFocus
 End If
    
End Sub
Public Sub Inicia(ByVal CNumProd As String, ByVal cAlcance As String)
     NumProd = CNumProd
     txtCuenta.Prod = NumProd
     txtCuenta.EnabledProd = False
        
End Sub

Private Sub txtDesde_Validate(Cancel As Boolean)

  If txtDesde.Text <> "" Then
    If Not (CLng(txtDesde.Text) >= CLng(lblRangoIni.Caption) And CLng(txtDesde.Text) < CLng(lblRangoIni.Caption)) Then
           txtDesde.SetFocus
           Cancel = True
    End If
  End If
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
 If KeyAscii <> 13 Then
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
      End If
 Else
          If txtHasta.Text <> "" Then
            If Not (CLng(txtHasta.Text) > CLng(lblRangoIni.Caption) And CLng(txtHasta.Text) <= CLng(lblRangoFin.Caption)) Then
                
                txtHasta.SetFocus
                Exit Sub
            End If
          End If
        cmdImprimir.SetFocus
 End If
End Sub

Private Sub txtHasta_Validate(Cancel As Boolean)

  If txtHasta.Text <> "" Then
    If Not (CLng(txtHasta.Text) > CLng(lblRangoIni.Caption) Or CLng(txtHasta.Text) <= CLng(lblRangoFin.Caption)) Then
             Cancel = True
           txtHasta.SetFocus
          
    End If
  End If
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
 If KeyAscii <> 13 Then
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
      End If
 Else
    If txtNum.Text <> "" Then
        If Not (CLng(txtNum.Text) >= CLng(lblRangoIni.Caption) And CLng(txtNum.Text) <= CLng(lblRangoFin.Caption)) Then
            txtNum.SetFocus
            Exit Sub
        End If
    End If
     Me.txtNum.SetFocus
 End If
End Sub

Private Sub txtNum_Validate(Cancel As Boolean)

  If txtNum.Text <> "" Then
    If Not (CLng(txtNum.Text) >= CLng(lblRangoIni.Caption) And CLng(txtNum.Text) <= CLng(lblRangoFin.Caption)) Then
            txtNum.SetFocus
           Cancel = True
    End If
  End If
End Sub
