VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImprimirCuponesBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Cupones de Campaña"
   ClientHeight    =   2820
   ClientLeft      =   2970
   ClientTop       =   4005
   ClientWidth     =   8130
   Icon            =   "frmImprimirCuponesBatch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   135
      TabIndex        =   2
      Top             =   2340
      Width           =   1080
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   6795
      TabIndex        =   1
      Top             =   2355
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   7845
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias..."
         Height          =   345
         Left            =   6525
         TabIndex        =   17
         Top             =   1035
         Width           =   1140
      End
      Begin VB.CheckBox chkImpresos 
         Caption         =   "Solo  Impresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2505
         TabIndex        =   12
         Top             =   1065
         Width           =   1890
      End
      Begin VB.CheckBox chkNoImpresos 
         Caption         =   "Solo no Impresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   11
         Top             =   1065
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   405
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtHasta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4695
         MaxLength       =   10
         TabIndex        =   9
         Top             =   510
         Width           =   1320
      End
      Begin VB.TextBox txtDesde 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   8
         Top             =   525
         Width           =   1320
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Rango "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   405
         TabIndex        =   5
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Left            =   4050
         TabIndex        =   7
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   1830
         TabIndex        =   6
         Top             =   585
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   120
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNroCta 
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
      Height          =   285
      Left            =   5565
      TabIndex        =   16
      Top             =   90
      Width           =   1425
   End
   Begin VB.Label lblRangoFin 
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
      Height          =   285
      Left            =   5565
      TabIndex        =   15
      Top             =   465
      Width           =   1425
   End
   Begin VB.Label Label5 
      Caption         =   "Rango Máximo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4185
      TabIndex        =   14
      Top             =   510
      Width           =   1410
   End
   Begin VB.Label Label4 
      Caption         =   "Nro Cuentas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4185
      TabIndex        =   13
      Top             =   165
      Width           =   1260
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
      Left            =   135
      TabIndex        =   4
      Top             =   210
      Width           =   645
      WordWrap        =   -1  'True
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
      Left            =   870
      TabIndex        =   3
      Top             =   150
      Width           =   2910
   End
End
Attribute VB_Name = "frmImprimirCuponesBatch"
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
Dim oSorteo As DSorteo, rsTemp As Recordset
    
    Set oSorteo = New DSorteo
    Set rsTemp = oSorteo.GetSorteo("I", gsAlcance)
    If Not rsTemp.EOF Then
          ReDim MatSorteos(rsTemp.RecordCount)
          i = 1
          While Not rsTemp.EOF
                  Me.lblNroCta.Caption = rsTemp!NroCuentas
                  Me.lblRangoFin.Caption = rsTemp!NroRangoMax
                  MatSorteos(i) = rsTemp!CnumSorteo
                      i = i + 1
                  rsTemp.MoveNext
                  If rsTemp.EOF Then i = i - 1
                  
          Wend
          
            NumSort = UBound(MatSorteos)
    End If
    
End Sub
Private Sub Limpiar()
    
    'lblTitular.Caption = ""
    Me.lblNroCta.Caption = ""
    Me.lblRangoFin.Caption = ""
    Me.txtDesde.Text = ""
    Me.txtHasta.Text = ""
    OptTipo(1).value = True
    chkNoImpresos.value = vbChecked
    
    
End Sub

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdImprimir_Click()
Dim oSorteo As DSorteo
Set oSorteo = New DSorteo

'On Error GoTo ErrMsg

If MsgBox("Este proceso podría tardar varios minutos." & vbCrLf & "Esta seguro de procesarlo.", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Screen.MousePointer = vbHourglass
    Dim lnAge As Integer, nRangoIni As String, nRangoFin As String, Nimpresiones As Long
    Dim oPrevio As Previo.clsPrevio, lscadena As String, rstmpCup As Recordset
    Dim orep As nCaptaReportes
    Set orep = New nCaptaReportes
    Set oPrevio = New Previo.clsPrevio
     lscadena = ""
    
    Set rstmpCup = New Recordset
    
    nRangoIni = "0": nRangoFin = "0"
    
     If chkImpresos.value = vbChecked And chkNoImpresos.value = vbUnchecked Then
         Nimpresiones = 1
     ElseIf chkImpresos.value = vbUnchecked And chkNoImpresos.value = vbChecked Then
         Nimpresiones = 0
     ElseIf (chkImpresos.value = vbChecked And chkNoImpresos.value = vbChecked) Or (chkImpresos.value = vbUnchecked And chkNoImpresos.value = vbUnchecked) Then
         Nimpresiones = -1
     End If
     
     If OptTipo(0).value = True Then
        If Trim(txtDesde.Text) = "" Or Trim(txtHasta.Text = "") Then
            MsgBox "Ingrese el rango correspondiente para la impresión", vbOKOnly + vbInformation, "Aviso"
            Exit Sub
        End If
        If CLng(Trim(txtDesde.Text)) > CLng(Trim(txtHasta.Text)) Then
            MsgBox "El valor del rango hasta debe ser mayor al valor del rango desde ", vbOKOnly + vbInformation, "Aviso"
            Exit Sub
        End If
        If CLng(Trim(txtDesde.Text)) <= CLng(Trim(txtHasta.Text)) Then
            nRangoIni = Trim(txtDesde.Text): nRangoFin = Trim(txtHasta.Text)
        End If
     End If
     
             
     If gsAlcance = "00" Then
       
            For lnAge = 1 To frmSelectAgencias.List1.ListCount
                 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                   lscadena = lscadena & orep.CUPONESBATCH(Left(Trim(Me.lblNumSorteo.Caption), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), , nRangoIni, nRangoFin, Nimpresiones, rstmpCup)
                        
                 End If
                 If rstmpCup.State = 1 Then
                    If Not rstmpCup.EOF Then
                        rstmpCup.MoveFirst
            
                        While Not rstmpCup.EOF
                            Call oSorteo.ActualizaCuentaImpresion(Trim(Me.lblNumSorteo.Caption), rstmpCup!cNumCta)
                            rstmpCup.MoveNext
                        Wend
                     End If
                End If
                 
            Next lnAge
        
       
     Else
         For lnAge = 1 To frmSelectAgencias.List1.ListCount
                
                If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                    Set rstmpCup = New Recordset
                    lscadena = lscadena & orep.CUPONESBATCH(Left(Trim(Me.lblNumSorteo.Caption), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), , nRangoIni, nRangoFin, Nimpresiones, rstmpCup)
                Else
                    Set rstmpCup = Nothing
                End If
                
                If Not (rstmpCup Is Nothing) Then
                
                    If Not rstmpCup.EOF Then
                        rstmpCup.MoveFirst
            
                        While Not rstmpCup.EOF
                            Call oSorteo.ActualizaCuentaImpresion(Trim(Me.lblNumSorteo.Caption), rstmpCup!cNumCta)
                            rstmpCup.MoveNext
                        Wend
                    End If
                End If
               
                
          Next lnAge
     End If
            
'        If Me.optImpresion(0).value = True Then
'            Set loPrevio = New Previo.clsPrevio
'                loPrevio.Show lscadena, "Cupones" & Trim(TxtNumSorteo.Text), True
'            Set loPrevio = Nothing
'        Else
                Set oPrevio = New Previo.clsPrevio
                    oPrevio.PrintSpool sLpt, lscadena, True
                Set oPrevio = Nothing
            
'            dlgGrabar.CancelError = True
'            dlgGrabar.InitDir = App.path
'            dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'            dlgGrabar.ShowSave
'            If dlgGrabar.FileName <> "" Then
'               Open dlgGrabar.FileName For Output As #1
'                Print #1, lscadena
'                Close #1
'            End If
'        End If
       Set oSorteo = Nothing
       Set rstmpCup = Nothing
       
       Screen.MousePointer = vbDefault
        
End If

'Exit Sub
'ErrMsg:
'MsgBox "Ocurrió el error:" & Err.Number & " " & Err.Description, vbOKOnly, "Aviso"
'  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()
 Unload Me
 
End Sub

Private Sub Form_Load()
  gsAlcance = frmPreparaSorteo.gsAlcance
    CargaMatriz (gsAlcance)
    If i > 0 Then
       lblNumSorteo.Caption = MatSorteos(i)
    End If
    Me.Left = 2925
    Me.Top = 3675
End Sub

Private Sub optTipo_Click(Index As Integer)
Select Case Index
   Case 0
         txtDesde.Text = ""
         txtHasta.Text = ""
         
         txtDesde.Enabled = True
         txtHasta.Enabled = True
         
         txtDesde.SetFocus
   Case 1
         txtDesde.Text = ""
         txtHasta.Text = ""
         
         txtDesde.Enabled = False
         txtHasta.Enabled = False
         
   
End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If OptTipo(0).value = True Then
        txtHasta.SetFocus
     End If
  ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
     KeyAscii = 0
  End If
End Sub
