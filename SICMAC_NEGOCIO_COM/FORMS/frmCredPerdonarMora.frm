VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCredPerdonarMora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perdonar Mora"
   ClientHeight    =   6675
   ClientLeft      =   600
   ClientTop       =   2115
   ClientWidth     =   11940
   Icon            =   "frmCredPerdonarMora.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5145
      Left            =   105
      TabIndex        =   5
      Top             =   1440
      Width           =   11725
      Begin VB.TextBox txtGlosa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   3840
         Width           =   6330
      End
      Begin VB.Frame fraPerdonMora 
         Caption         =   "Perdón de Mora"
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
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   4935
         Begin VB.TextBox txtPorcentaje 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Text            =   "100.00"
            Top             =   315
            Width           =   855
         End
         Begin MSComCtl2.UpDown udPorcentaje 
            Height          =   255
            Left            =   2640
            TabIndex        =   16
            Top             =   315
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label lblPorc 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   2880
            TabIndex        =   15
            Top             =   360
            Width           =   120
         End
         Begin VB.Label lblMontoNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3720
            TabIndex        =   14
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label lblMonto 
            AutoSize        =   -1  'True
            Caption         =   "Monto:"
            Height          =   195
            Left            =   3120
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblPerdonMora 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje a Perdonar:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   1350
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   10200
         TabIndex        =   8
         Top             =   4560
         Width           =   1350
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8760
         TabIndex        =   7
         Top             =   4560
         Width           =   1350
      End
      Begin SICMACT.FlexEdit FECalend 
         Height          =   3450
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   6085
         Cols0           =   17
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCredPerdonarMora.frx":030A
         EncabezadosAnchos=   "400-1000-400-1000-1000-1000-1000-1100-1400-1100-1400-1000-1200-0-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-2-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
      Begin VB.Frame fraPerdonMoraCamp 
         Caption         =   "Perdón de Mora"
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
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Visible         =   0   'False
         Width           =   6735
         Begin VB.ComboBox cmbModalidad 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   300
            Width           =   2820
         End
         Begin VB.Label lblMontoCamp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5520
            TabIndex        =   24
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label lblModalidad 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblPorPerdon 
            AutoSize        =   -1  'True
            Caption         =   "%Perdon:"
            Height          =   195
            Left            =   3960
            TabIndex        =   22
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lblPorcCamp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4680
            TabIndex        =   21
            Top             =   315
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Credito"
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
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11685
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   7845
         TabIndex        =   2
         Top             =   540
         Width           =   1200
      End
      Begin VB.Frame fraCamp 
         Caption         =   "Campaña de Recuperaciones"
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
         Height          =   615
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         Begin VB.ComboBox cmbCampRec 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   2820
         End
      End
      Begin VB.CheckBox ChkPerdMora 
         Caption         =   "Perdonar Dias de Atraso"
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   9195
         TabIndex        =   3
         Top             =   195
         Width           =   2310
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredPerdonarMora.frx":03A5
            Left            =   100
            List            =   "frmCredPerdonarMora.frx":03A7
            TabIndex        =   4
            Top             =   225
            Width           =   2115
         End
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   195
         TabIndex        =   1
         Top             =   495
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredPerdonarMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto

'Se agrego para mantener el esquema de Componentes
Dim nNroCalen As Integer
Dim nNroCuota As Integer
Dim objPista As COMManejador.Pista
'WIOR 20130907 ********************
Event KeyPress(KeyAscii As Integer)
Private fsIntMora As String
'WIOR FIN **************************
'WIOR 20150331 ********************
Dim fnActiva As Integer
Dim MatModalidades As Variant
Dim MatPerdonCamp As Variant
Dim fbValidaCampRecup As Boolean
Dim fnCuotaIniVenc As Integer
Dim fnCantCuotasMora As Integer
Dim fnDiasAtraso As Integer
'WIOR FIN *************************

Public Sub LimpiaPantalla()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LimpiaFlex FECalend
    'WIOR 20130907 ********************
    CmdAceptar.Enabled = False
    fraPerdonMora.Enabled = False
    LstCred.Clear 'NAGL 20190815
    txtGlosa.Locked = True 'NAGL 20190815 Según ERS036-2018
    txtGlosa.Text = "" 'NAGL 20190815
    txtPorcentaje.Text = "100.00"
    lblMontoNeto.Caption = ""
    fsIntMora = ""
    'WIOR FIN *************************
    'WIOR 20150401 ***************
    fnCuotaIniVenc = 0
    fnDiasAtraso = 0
    fnCantCuotasMora = 0
    lblPorcCamp.Caption = ""
    lblMontoCamp.Caption = ""
    'WIOR FIN ********************
    If fnActiva = 1 Then
        If Trim(Right(cmbCampRec.Text, 2)) <> "1" Then
           cmbModalidad.Clear
           fraPerdonMoraCamp.Enabled = False
        End If
    End If
    'WIOR FIN *******************
End Sub

Private Function HabilitaActualizacion(ByVal pbHabilita As Boolean) As Boolean
    Frame3.Enabled = Not pbHabilita
    'CmdAceptar.Enabled = pbHabilita'WIOR 20130909 COMENTO
    CmdCancelar.Enabled = pbHabilita
    CmdSalir.Enabled = Not pbHabilita
End Function

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim oCalend As COMDCredito.DCOMCalendario
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim nMontoApr As Double
Dim nTasaInteres As Double
Dim nEstado As Integer
'WIOR 20150331 ***********
Dim dFechaVenc As Date
Dim nContador As Integer
nContador = 0
'WIOR FIN ****************
    
    'WIOR 20150331 **************
    If Not ValidaCampRecup Then
        fbValidaCampRecup = True
        Exit Function
    End If
    'WIOR FIN *******************
    
    'WIOR 20130907 ********************
    CmdAceptar.Enabled = False
    fraPerdonMora.Enabled = False
    txtGlosa.Locked = True 'NAGL 20190815 Según ERS036-2018
    txtGlosa.Text = "" 'NAGL 20190815
    txtPorcentaje.Text = "100.00"
    lblMontoNeto.Caption = ""
    fsIntMora = ""
    'WIOR FIN *************************
    On Error GoTo ErrorCargaDatos
    LimpiaFlex FECalend
    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.CargaDatosPerdonarMora(psCtaCod, nMontoApr, nTasaInteres, R2, R, nNroCalen, nEstado, fnDiasAtraso, gsPerdonMoraCred) 'WIOR 20150331 AGREGO fnDiasAtraso 'NAGL Agregó 20190816 Según ERS036-2018
    Set oCredito = Nothing
'    Set oCalend = New COMDCredito.DCOMCalendario
'    Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
'    Set oCalend = Nothing
    
'    Set oCredito = New COMDCredito.DCOMCredito
'    Set R2 = oCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
'    If R2.BOF Or R2.EOF Then
'        nMontoApr = 0
'    Else
'        nMontoApr = CDbl(Format(IIf(IsNull(R2!nMonto), 0, R2!nMonto), "#0.00"))
'    End If
'    R2.Close
'    Set R2 = Nothing
    
'    Set R2 = oCredito.RecuperaProducto(psCtaCod)
'    nTasaInteres = CDbl(Format(R2!nTasaInteres, "#0.00"))
'    R2.Close
'    Set R2 = Nothing
'    Set oCredito = Nothing
    
    Set R2 = Nothing
    
    If R.BOF Or R.EOF Then
        CargaDatos = False
        Exit Function
    Else
        CargaDatos = True
    End If
    
    If nEstado <> gColocEstVigNorm And nEstado <> gColocEstVigVenc And nEstado <> gColocEstVigMor _
       And nEstado <> gColocEstRefNorm And nEstado <> gColocEstRefVenc And nEstado <> gColocEstRefMor _
       And gColocEstRecVigJud <> 2205 Then
        CargaDatos = False
        Exit Function
    End If
    
    FECalend.ForeColor = vbBlack 'WIOR 20150401
    Do While Not R.EOF
        FECalend.AdicionaFila
        FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dVenc, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 2) = Trim(str(R!nCuota))
        FECalend.TextMatrix(R.Bookmark, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital) + _
                                        IIf(IsNull(R!nIntComp), 0, R!nIntComp) + _
                                        IIf(IsNull(R!nIntGracia), 0, R!nIntGracia) + _
                                        IIf(IsNull(R!nIntMor), 0, R!nIntMor) + _
                                        IIf(IsNull(R!nIntReprog), 0, R!nIntReprog) + _
                                        IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso) + _
                                        IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 4) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 5) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 6) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00") 'Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        'FECalend.TextMatrix(R.Bookmark, 7) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor) + IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00") 'Format(IIf(IsNull(R!nIntMorPag), 0, R!nIntMorPag), "#0.00") 'Comentado by NAGL20190717
        FECalend.TextMatrix(R.Bookmark, 7) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00") 'NAGL ERS036-2018 20190717
        'FECalend.TextMatrix(R.Bookmark, 8) = Format(IIf(IsNull(R!nIntMorPag), 0, R!nIntMorPag) + IIf(IsNull(R!nIntCompVencPag), 0, R!nIntCompVencPag), "#0.00") 'Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")'Comentado by NAGL20190717
        FECalend.TextMatrix(R.Bookmark, 8) = Format(IIf(IsNull(R!nInMorPerdonada), 0, R!nInMorPerdonada), "#0.00") 'NAGL ERS036-2018 20190717
        FECalend.TextMatrix(R.Bookmark, 9) = Format(IIf(IsNull(R!nIntMorPagada), 0, R!nIntMorPagada), "#0.00") 'NAGL ERS036-2018 20190717
        FECalend.TextMatrix(R.Bookmark, 10) = Format(FECalend.TextMatrix(R.Bookmark, 7) - FECalend.TextMatrix(R.Bookmark, 8) - FECalend.TextMatrix(R.Bookmark, 9), "#0.00") 'Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        'NAGL Agregó - FECalend.TextMatrix(R.Bookmark, 9) Según ERS036-2018 20190816
        nMontoApr = nMontoApr - IIf(IsNull(R!nCapital), 0, R!nCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.00"))
        FECalend.TextMatrix(R.Bookmark, 11) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00") 'Format(nMontoApr, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 12) = Format(nMontoApr, "#0.00") 'Trim(Str(R!nColocCalendEstado))
        
        '06-05-2005
        FECalend.TextMatrix(R.Bookmark, 13) = Trim(str(R!nColocCalendEstado)) 'Format(IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 14) = Format(IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 15) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 16) = nEstado 'NAGL ERS036-2018 20190717
        '***************************************
        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.row = R.Bookmark
            Call FECalend.ForeColorRow(vbRed)
        'WIOR 20150331 *****************
        ElseIf R!nInMorPerdonada > 0 Then 'NAGL 20190816
            FECalend.row = R.Bookmark
            Call FECalend.ForeColorRow(vbBlue)
        ElseIf R!nIntMorPagada > 0 Then 'NAGL 20190816
            FECalend.row = R.Bookmark
            Call FECalend.ForeColorRow(RGB(0, 114, 0))
        Else
            nContador = nContador + 1
            If nContador = 1 Then
                dFechaVenc = CDate(R!dVenc)
                fnCuotaIniVenc = CInt(R!nCuota)
            End If
            
            If CDbl(FECalend.TextMatrix(R.Bookmark, 7)) > 0 Then
                fnCantCuotasMora = fnCantCuotasMora + 1
            End If
        'WIOR FIN **********************
        End If
        
        R.MoveNext
    Loop
    R.Close
    FECalend.TopRow = 1 'WIOR 20130907
    
    Set R = Nothing
    'WIOR 20150331 **************
    If fnActiva = 1 Then
        If Trim(Right(cmbCampRec.Text, 2)) <> "1" Then
            If Not CargarModalidadesCampRecup(fnDiasAtraso, dFechaVenc) Then
                CargaDatos = False
                fbValidaCampRecup = True
                Call cmdCancelar_Click
            End If
        End If
    End If
    'WIOR FIN *******************
    Exit Function
ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub ActxCta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sCuenta As String
Dim bRetSinTarjeta As Boolean

    If KeyCode = 123 Then ' cuando se pulsa la tecla F12
        sCuenta = frmValTarCodAnt.inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
     End If
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorActxCta_KeyPress
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            HabilitaActualizacion False
            If Not fbValidaCampRecup Then 'WIOR 20150401
                MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
            End If 'WIOR 20150401
        Else
            HabilitaActualizacion True
        End If
    End If
    Exit Sub

ErrorActxCta_KeyPress:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

'WIOR 20150330 ***********************
Private Sub cmbCampRec_Click()
    If Trim(Right(cmbCampRec.Text, 2)) = "1" Then
        fraPerdonMora.Visible = True
        fraPerdonMoraCamp.Visible = False
        cmbModalidad.Clear
        Me.ChkPerdMora.Visible = True
        FECalend.HighLight = flexHighlightAlways
    Else
        fraPerdonMora.Visible = False
        fraPerdonMoraCamp.Visible = True
        ChkPerdMora.Visible = False
        FECalend.HighLight = flexHighlightNever
    End If
End Sub
'WIOR FIN ***************************

'WIOR 20150401 *****************
Private Sub cmbModalidad_Click()
Dim i As Integer
Dim nIndex As Integer
Dim nMontoPerdon As Double

For i = 1 To UBound(MatModalidades, 2)
    If Trim(MatModalidades(0, i)) = Trim(Right(cmbModalidad.Text, 3)) Then
        nIndex = i
        Exit For
    End If
Next i

If (fnCantCuotasMora - 1) < CInt(MatModalidades(2, nIndex)) Then
    MsgBox "Solo cuenta con " & fnCantCuotasMora & IIf(fnCantCuotasMora = 1, " cuota vencida.", " cuotas vencidas."), vbInformation, "Aviso"
    cmbModalidad.ListIndex = -1
    Exit Sub
End If

ReDim MatPerdonCamp(4, 0 To ((fnCantCuotasMora - CInt(MatModalidades(2, nIndex))) - 1))
nMontoPerdon = 0
For i = 0 To ((fnCantCuotasMora - CInt(MatModalidades(2, nIndex))) - 1)
    MatPerdonCamp(0, i) = fnCuotaIniVenc + i
    MatPerdonCamp(1, i) = FECalend.TextMatrix(fnCuotaIniVenc + i, 9)
    MatPerdonCamp(2, i) = MatModalidades(3, nIndex)
    MatPerdonCamp(3, i) = Round(FECalend.TextMatrix(fnCuotaIniVenc + i, 9) * CDbl(MatModalidades(3, nIndex)), 2)
    nMontoPerdon = nMontoPerdon + CDbl(MatPerdonCamp(3, i))
Next i

lblPorcCamp.Caption = Format(MatModalidades(3, nIndex) * 100, "##0.00")
lblMontoCamp.Caption = Format(nMontoPerdon, "##0.00")
End Sub
'WIOR FIN **********************

Private Sub CmdAceptar_Click()
    Dim oCal As COMDCredito.DCOMCredActBD
    Dim loCred As COMDCredito.DCOMCredito
    Dim lrsCred As ADODB.Recordset
    Dim sMovNro As String
    Dim nMovNro As Long
    Dim nMontoPerdonar As Double 'WIOR 20130909
    Dim nNroCuotaPerd As Integer 'NAGL 20190715
    Dim nPrdEstado As Integer
    '****Agregado by NAGL 20190715***
    Dim lrsOperCond As ADODB.Recordset
    Dim lsCadena As String
    Dim pnMora As Double
    Dim pnMoraPagadaAnt As Double
    Dim pnMoraAPerdonar As Double
    Dim pnMoraAPerdonarPorc As Double
    Dim pnMoraPerdon As Double
    Dim pnMoraPagada As Double
    Dim psGlosa As String
    '******END NAGL 20190717*********

    On Error GoTo ErrorCmdAceptar_Click
    
    'WIOR 20150401 ****************
    Dim bCampRecup As Boolean
    bCampRecup = False
    
    If fnActiva = 1 Then
        If Trim(Right(cmbCampRec.Text, 2)) <> "1" Then
            bCampRecup = True
        End If
    End If
        
    Dim Msj As String
    Msj = gVarPublicas.ValidarFechaSistServer
    If Msj <> "" Then
        MsgBox Msj, vbInformation, "Aviso"
        Unload frmCredPerdonarMora
    Else
        If bCampRecup Then
            Call GrabarPerdonMoraCampRecup
        Else
        'WIOR FIN *********************
            If CInt(FECalend.TextMatrix(FECalend.row, 13)) = gColocCalendEstadoPagado Then
                MsgBox "No Puede Perdonar la Mora de una Cuota Cancelada", vbInformation, "Aviso"
                Exit Sub
            End If
            
            '*****NAGL 20190719 Según ERS036-2018***********
            nNroCuotaPerd = CInt(FECalend.TextMatrix(FECalend.row, 2))
            
            If CDbl(FECalend.TextMatrix(FECalend.row, 7)) = 0 Then
                MsgBox "No se puede perdonar la Mora de una Cuota con monto cero", vbInformation, "Aviso"
                Exit Sub
            End If
            
            If CDbl(FECalend.TextMatrix(FECalend.row, 10)) = 0 Then
                MsgBox "La Mora ya se encuentra saldada, por favor intente con la siguiente cuota", vbInformation, "Aviso"
                Exit Sub
            End If
            
            Set loCred = New COMDCredito.DCOMCredito
            If loCred.esAplicableCuotaAPerdonar(ActxCta.NroCuenta, nNroCuotaPerd) = False Then
                MsgBox "No se puede perdonar la Mora de esta Cuota, debe considerar la más próxima...", vbInformation, "Aviso"
                Exit Sub
            End If
            Set loCred = Nothing
            
            If txtGlosa = "" Then
                MsgBox "Por favor ingrese la glosa correspondiente", vbInformation, "Aviso"
                Exit Sub
            End If
            '-------------Mensaje de Advertencia-----------
            Set lrsOperCond = New ADODB.Recordset
            Set loCred = New COMDCredito.DCOMCredito
            Set lrsOperCond = loCred.ObtieneHistorialCuotasCond(ActxCta.NroCuenta, gdFecSis)
            If lrsOperCond!NroCuotas > 0 Then
               lsCadena = "En el historial se tiene la siguiente condonación: "
               Do While Not lrsOperCond.EOF
                  lsCadena = lsCadena & " " & lrsOperCond!cOpeDescGeneral + " , "
                  lrsOperCond.MoveNext
               Loop
               lsCadena = Mid(lsCadena, 1, Len(lsCadena) - 3)
               'MsgBox lsCadena, vbInformation, "Aviso"
               If MsgBox(lsCadena & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    Exit Sub
               End If
            End If
            
            Set lrsOperCond = Nothing
            Set loCred = Nothing
            '-----------------------------------------------
            '********END NAGL 20190716**********************
            
            If MsgBox("Se va a Perdonar la Mora, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
            
            '*******NAGL 20190716**********
            CmdAceptar.Caption = "Procesando"
            CmdAceptar.Enabled = False
            CmdCancelar.Enabled = False
            '********************************
            
            '*DAOR 20070228, Obtener Dias de Atraso ************************
            Set loCred = New COMDCredito.DCOMCredito
            Set lrsCred = loCred.RecuperaColocacCredCampos(ActxCta.NroCuenta, "nDiasAtraso")
            Set loCred = Nothing
            '***************************************************************
            
            Set oCal = New COMDCredito.DCOMCredActBD
            
            If Me.ChkPerdMora.value = 1 Then
                Call oCal.dUpdateColocacCred(ActxCta.NroCuenta, 0)
            End If
            nMontoPerdonar = CDbl(lblMontoNeto.Caption) 'WIOR 20130909
            'nMontoPerdonar = nMontoPerdonar + CDbl(FECalend.TextMatrix(FECalend.row, 8)) + CDbl(FECalend.TextMatrix(FECalend.row, 9)) 'NAGL 20190716
            psGlosa = txtGlosa
            'Call oCal.dUpdateColocCalendDet(ActxCta.NroCuenta, nNroCalen, gColocCalendAplCuota, CInt(FECalend.TextMatrix(FECalend.Row, 2)), gColocConceptoCodInteresMoratorio, , CDbl(FECalend.TextMatrix(FECalend.Row, 14)), , False)'WIOR 20130909 COMENTÓ
            Call oCal.dUpdateColocCalendDet(ActxCta.NroCuenta, nNroCalen, gColocCalendAplCuota, CInt(FECalend.TextMatrix(FECalend.row, 2)), gColocConceptoCodInteresMoratorio, , nMontoPerdonar, , False, True) 'WIOR 20130909
            'NAGL 20191217 Agregó True al Final
            '05-05-2005
            'Call oCal.dUpdateColocCalendDet(ActxCta.NroCuenta, nNroCalen, gColocCalendAplCuota, CInt(FECalend.TextMatrix(FECalend.row, 2)), gColocConceptoCodInteresCompVencido, , CDbl(FECalend.TextMatrix(FECalend.row, 13)), , False) 'Comentado by NAGL 20190719 Según ERS036-2018, para no considerar Int.Compensatorio Vencido
            '******************************
            
            'Insertando movimiento
            sMovNro = oCal.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            'Call oCal.dInsertMov(sMovNro, "100904", "PERDONAR MORA", gMovEstContabMovContable, gMovFlagExtornado, False) 'Comentado by NAGL Según ERS036-2018
            Call oCal.dInsertMov(sMovNro, gsPerdonMoraCred, "PERDONAR MORA", gMovEstContabMovContable, gMovFlagVigente, False) 'NAGL 20190715 Según ERS036-2018
            nMovNro = oCal.dGetnMovNro(sMovNro)
           
            '*****Agregado by NAGL 20190715 Según ERS036-2019********
            Set lrsOperCond = New ADODB.Recordset
            Set loCred = New COMDCredito.DCOMCredito
            Set lrsOperCond = loCred.GetOpeCondonacionxEstado(ActxCta.NroCuenta, "cOpeCond")
            Set loCred = Nothing
            nPrdEstado = CInt(FECalend.TextMatrix(FECalend.row, 16))
             
            'Traslado a esta sección by NAGL 20190718
            Call oCal.dInsertMovCol(nMovNro, lrsOperCond("cOpeCond"), ActxCta.NroCuenta, nNroCalen, CDbl(lblMontoNeto.Caption), lrsCred("nDiasAtraso"), "GMIC", 0, 0, nPrdEstado) 'WIOR 20130909
            'NAGL 20190716 Cambió de nMontoPerdonar a CDbl(lblMontoNeto.Caption)
            'NAGL 20190716 Cambió de "100904" a lrsOperCond("cOpeCond")
            'NAGL 20190717 Agregó nPrdEstado
            Call oCal.dInsertMovColDet(nMovNro, lrsOperCond("cOpeCond"), ActxCta.NroCuenta, nNroCalen, gColRecConceptoCodInteresMoratorio, nNroCuotaPerd, CDbl(lblMontoNeto.Caption))
            
            '---Para Agregar en CredPerdonMora como Historial
            Set loCred = New COMDCredito.DCOMCredito
            pnMora = CDbl(FECalend.TextMatrix(FECalend.row, 7))
            pnMoraPagadaAnt = CDbl(FECalend.TextMatrix(FECalend.row, 8)) + CDbl(FECalend.TextMatrix(FECalend.row, 9)) 'NAGL 20191221
            pnMoraAPerdonar = CDbl(FECalend.TextMatrix(FECalend.row, 10))
            pnMoraAPerdonarPorc = CDbl(txtPorcentaje.Text) / 100
            pnMoraPerdon = CDbl(lblMontoNeto.Caption)
            pnMoraPagada = pnMoraPagadaAnt + nMontoPerdonar 'Mora Pagada Final
            
            Call loCred.AgregaRegCredPerdonMora(nMovNro, ActxCta.NroCuenta, nNroCalen, nNroCuotaPerd, pnMora, pnMoraPagadaAnt, pnMoraAPerdonar, pnMoraAPerdonarPorc, pnMoraPerdon, pnMoraPagada, sMovNro, gsPerdonMoraCred, psGlosa, True)
            Set loCred = Nothing
            '****************END NAGL 20190716**********************
            
            ''*** PEAC 20090220
             objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Perdonar mora", ActxCta.NroCuenta, gCodigoCuenta
            
            Set oCal = Nothing
            Set lrsCred = Nothing
            If MsgBox("Desea perdonar mora de otra cuota?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes Then
                Call CargaDatos(ActxCta.NroCuenta)
                '*******NAGL 20190716**********
                CmdAceptar.Caption = "Aceptar"
                '********************************
                Exit Sub
            End If
            '*******NAGL 20190716***********
            CmdAceptar.Caption = "Aceptar"
            '********************************
            Call cmdCancelar_Click
            
        End If 'WIOR 20150401
    End If      ''FIN ANGC 20200306
    Exit Sub
ErrorCmdAceptar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona

    On Error GoTo ErrorCmdBuscar_Click
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        'Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigNorm)) 'Comentado by NAGL 20190719
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigNorm, gColocEstVigVenc, gColocEstVigMor, gColocEstRefNorm, gColocEstRefVenc, gColocEstRefMor)) 'NAGL 20190719 Según ERS036-2018, Sección Estados del Producto
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
    Exit Sub

ErrorCmdBuscar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
    HabilitaActualizacion False
    LimpiaPantalla
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'WIOR 20130907 *******************
Private Sub FECalend_Click()
    If fnActiva = 0 Or Trim(Right(cmbCampRec.Text, 3)) = "1" Then 'WIOR 20150401
        CmdAceptar.Enabled = True
        fraPerdonMora.Enabled = True
        txtGlosa.Locked = False 'NAGL 20190815 Según ERS036-2018
        'fsIntMora = FECalend.TextMatrix(Me.FECalend.row, 14) 'Comentado by NAGL 20190716
        fsIntMora = FECalend.TextMatrix(Me.FECalend.row, 10) 'NAGL 20190716 Según ERS036-2018
        txtPorcentaje.Text = "100.00" 'NAGL 20190716
        Call CalcularPerdon
    End If 'WIOR 20150401
End Sub
'WIOR FIN ************************

Private Sub Form_Load()
    
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredPerdonarMora
    'WIOR 20130907 ********************
    CmdAceptar.Enabled = False
    fraPerdonMora.Enabled = False
    txtGlosa.Locked = True 'NAGL 20190815 Según ERS036-2018
    txtGlosa.Text = "" 'NAGL 20190815
    'WIOR FIN *************************
    'WIOR 20150330 ********************
    Call CargarCampanaRecup
    'WIOR FIN *************************
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub LstCred_Click()
    If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
        ActxCta.NroCuenta = LstCred.Text
    End If
End Sub

'WIOR 20130907 **********************************************
Private Sub txtPorcentaje_Change()
If IsNumeric(txtPorcentaje.Text) Then
    If Trim(txtPorcentaje.Text) = "" Or Trim(txtPorcentaje.Text) = "." Then
        txtPorcentaje.Text = "0.00"
    End If
    
    If CDbl(txtPorcentaje.Text) < 0 Then
        txtPorcentaje.Text = "0.00"
    End If
    
    If CDbl(txtPorcentaje.Text) > 100 Then
        txtPorcentaje.Text = "100.00"
    End If
Else
    txtPorcentaje.Text = "0.00"
End If

Call CalcularPerdon
End Sub

Private Sub txtPorcentaje_GotFocus()
    txtPorcentaje.SelStart = 0
    txtPorcentaje.SelLength = Len(txtPorcentaje)
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
Dim lsCadeNum As String
lsCadeNum = "0123456789."
RaiseEvent KeyPress(KeyAscii)
If InStr(1, lsCadeNum, Chr(KeyAscii), vbTextCompare) = 0 Then
    KeyAscii = 0
    CmdAceptar.SetFocus 'NAGL 20190715
End If
End Sub

Private Sub txtPorcentaje_LostFocus()
 txtPorcentaje = IIf(txtPorcentaje = "", 0, txtPorcentaje)
    If IsNumeric(txtPorcentaje) = False Then txtPorcentaje = 0
    
    If CDbl(txtPorcentaje.Text) < 0 Then
        txtPorcentaje.Text = 0
        Exit Sub
    End If
    If CDbl(txtPorcentaje.Text) > 100 Then
        txtPorcentaje.Text = 100
        Exit Sub
    End If
    txtPorcentaje.Text = Format(txtPorcentaje.Text, "##0.00")
End Sub

Private Sub udPorcentaje_DownClick()
Dim valor As Double
valor = CDbl(txtPorcentaje.Text) - 0.01
If valor < 0 Then
    valor = 0
End If
txtPorcentaje.Text = Format(valor, "##0.00")
End Sub

Private Sub udPorcentaje_UpClick()
Dim valor As Double
valor = CDbl(txtPorcentaje.Text) + 0.01
If valor > 100 Then
    valor = 100
End If
txtPorcentaje.Text = Format(valor, "##0.00")
End Sub

Private Sub CalcularPerdon()
If ValidaDatos Then
    Dim nPorcentaje As Double
    nPorcentaje = CDbl(txtPorcentaje.Text)
    lblMontoNeto.Caption = Format(CDbl(fsIntMora) * (nPorcentaje / 100), "##0.00")
End If
End Sub
Private Function ValidaDatos() As Boolean
ValidaDatos = True

If Trim(fsIntMora) = "" Then
    MsgBox "Seleccione la Cuota a Perdonar", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

End Function
'WIOR FIN ***************************************************

'WIOR 20150331 *********************
Private Sub CargarCampanaRecup()
Dim oConsSist As COMDConstSistema.NCOMConstSistema
Dim oCons As COMDConstantes.DCOMConstantes
Set oConsSist = New COMDConstSistema.NCOMConstSistema

fnActiva = CInt(oConsSist.LeeConstSistema(502))
fbValidaCampRecup = False
fnCuotaIniVenc = 0
fnCantCuotasMora = 0
fnDiasAtraso = 0

If fnActiva = 1 Then
    fraCamp.Visible = True
    fraPerdonMoraCamp.Enabled = False
   
    Set oCons = New COMDConstantes.DCOMConstantes
    Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(7097), cmbCampRec)
    Set oCons = Nothing
End If

Set oConsSist = Nothing
End Sub

Private Function CargarModalidadesCampRecup(ByVal pnDiasAtraso, ByVal psFechaVenc As Date) As Boolean
Dim nCampRecup As Integer
Dim nDiasAtrasoNew As Integer
Dim dFecActCamp As Date
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsDatos As ADODB.Recordset
Dim i As Integer

CargarModalidadesCampRecup = True
ReDim MatModalidades(4, 0)
nCampRecup = CInt(Trim(Right(cmbCampRec.Text, 2)))
Set oDCredito = New COMDCredito.DCOMCredito
Set rsDatos = oDCredito.ObtenerSolPerdonMoraCampRecupCred(Trim(ActxCta.NroCuenta), nCampRecup)

If Not (rsDatos.EOF And rsDatos.BOF) Then
    MsgBox "El crédito ya realizo el perdon de mora con la Campaña ''" & Trim(Mid(cmbCampRec.Text, 1, Len(cmbCampRec.Text) - 3)) & "''", vbInformation, "Aviso"
    CargarModalidadesCampRecup = False
    Set rsDatos = Nothing
    Exit Function
Else
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.ObtenerConfigCampRecup(nCampRecup)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        dFecActCamp = CDate(rsDatos!dFechaActivacion)
    Else
        dFecActCamp = gdFecSis
    End If
    
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.ObtenerModalidadesCampRecup(nCampRecup, pnDiasAtraso)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        ReDim MatModalidades(4, 0 To rsDatos.RecordCount)
        cmbModalidad.Clear
        For i = 1 To rsDatos.RecordCount
            MatModalidades(0, i) = rsDatos!nId
            MatModalidades(1, i) = Trim(rsDatos!CDescripcion)
            MatModalidades(2, i) = rsDatos!nCuotasDejada
            MatModalidades(3, i) = rsDatos!nPerdon
            cmbModalidad.AddItem Trim(rsDatos!CDescripcion) & Space(100) & Trim(str(rsDatos!nId))
            rsDatos.MoveNext
        Next i
        fraPerdonMoraCamp.Enabled = True
        CmdAceptar.Enabled = True
    Else
        Set rsDatos = Nothing
        nDiasAtrasoNew = DateDiff("D", psFechaVenc, dFecActCamp)
        
        Set rsDatos = oDCredito.ObtenerModalidadesCampRecup(nCampRecup, nDiasAtrasoNew)
        
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            ReDim MatModalidades(4, 0 To rsDatos.RecordCount)
            cmbModalidad.Clear
            For i = 1 To rsDatos.RecordCount
                MatModalidades(0, i) = rsDatos!nId
                MatModalidades(1, i) = Trim(rsDatos!CDescripcion)
                MatModalidades(2, i) = rsDatos!nCuotasDejada
                MatModalidades(3, i) = rsDatos!nPerdon
                cmbModalidad.AddItem Trim(rsDatos!CDescripcion) & Space(100) & Trim(str(rsDatos!nId))
                rsDatos.MoveNext
            Next i
            fraPerdonMoraCamp.Enabled = True
            CmdAceptar.Enabled = True
        Else
            MsgBox "No existen modalidades de Perdon de Mora para esta crédito con la Campaña ''" & Trim(Mid(cmbCampRec.Text, 1, Len(cmbCampRec.Text) - 3)) & "''", vbInformation, "Aviso"
            CargarModalidadesCampRecup = False
            Set rsDatos = Nothing
            Exit Function
        End If
        
    End If
End If

Set rsDatos = Nothing
Set oDCredito = Nothing
End Function
Private Function ValidaCampRecup() As Boolean
ValidaCampRecup = True
    If fnActiva = 1 Then
        If Trim(cmbCampRec.Text) = "" Then
            MsgBox "Favor de Selecionar un Campaña de Recuperaciones", vbInformation, "Aviso"
            ValidaCampRecup = False
            cmbCampRec.SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub GrabarPerdonMoraCampRecup()
Dim oDCredito As COMDCredito.DCOMCredito
Dim nId As Double
Dim nIndex As Integer
Dim i As Integer

Dim oCal As COMDCredito.DCOMCredActBD
Dim sMovNro As String
Dim nMovNro As Long

    
On Error GoTo ErrorGrabar
    
    
If Trim(Me.cmbModalidad.Text) = "" Then
    MsgBox "Favor de Selecionar un Modalidad de la Campaña de Recuperaciones", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("Se va a Grabar el Perdon de la Mora, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

Set oDCredito = New COMDCredito.DCOMCredito



For i = 1 To UBound(MatModalidades, 2)
    If Trim(MatModalidades(0, i)) = Trim(Right(cmbModalidad.Text, 3)) Then
        nIndex = i
        Exit For
    End If
Next i
Set oCal = New COMDCredito.DCOMCredActBD
'oDCredito.BeginTrans
sMovNro = oCal.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

nId = oDCredito.GrabarPerdonMoraCampRecup(Trim(ActxCta.NroCuenta), nNroCalen, CInt(Trim(Right(cmbCampRec.Text, 3))), CInt(Trim(Right(cmbModalidad.Text, 3))), _
                                          MatModalidades(2, nIndex), CDbl(MatModalidades(3, nIndex)), CDbl(lblMontoCamp.Caption), _
                                          sMovNro, gdFecSis)
If nId = 0 Then
    MsgBox "Ocurrió un error al momento de Grabar, Favor de comunicarte con el Departamento de TI.", vbInformation, "Aviso"
    oDCredito.RollbackTrans
    Exit Sub
End If

For i = 0 To UBound(MatPerdonCamp, 2)
    Call oDCredito.GrabarPerdonMoraCampRecupDet(nId, CInt(MatPerdonCamp(0, i)), CDbl(MatPerdonCamp(1, i)), CDbl(MatPerdonCamp(3, i)))
    Call oCal.dUpdateColocCalendDet(Trim(ActxCta.NroCuenta), nNroCalen, gColocCalendAplCuota, CInt(MatPerdonCamp(0, i)), gColocConceptoCodInteresMoratorio, (-1) * CDbl(MatPerdonCamp(3, i)), , , , , True)
Next i

Call oDCredito.ActualizaMontosPerdonMoraCampRecup(nId)
objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Perdonar mora - Campaña Recuperaciones", ActxCta.NroCuenta, gCodigoCuenta

'Insertando movimiento
Call oCal.dInsertMov(sMovNro, "100904", "PERDONAR MORA", gMovEstContabMovContable, gMovFlagExtornado, False)
nMovNro = oCal.dGetnMovNro(sMovNro)

'Se incluye el NroCalen, monto a perdonar y dias de atraso
Call oCal.dInsertMovCol(nMovNro, "100904", Trim(ActxCta.NroCuenta), nNroCalen, CDbl(lblMontoCamp.Caption), fnDiasAtraso, "GMIC", 0, 0, 0) 'WIOR 20130909
        
'oDCredito.CommitTrans

Set oDCredito = Nothing
MsgBox "Se grabó correctamente el Perdon de la Mora", vbInformation, "Aviso"
Call cmdCancelar_Click

Exit Sub
ErrorGrabar:
    'oDCredito.RollbackTrans
 MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'WIOR FIN **************************


