VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCajaChicaArqueo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueo de Caja Chica"
   ClientHeight    =   7515
   ClientLeft      =   1740
   ClientTop       =   855
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaChicaArqueo.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMovDesc 
      Height          =   480
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6945
      Width           =   5265
   End
   Begin Sicmact.Usuario usu 
      Left            =   90
      Top             =   6840
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   6750
      TabIndex        =   11
      Top             =   7020
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   5520
      TabIndex        =   10
      Top             =   7020
      Width           =   1245
   End
   Begin VB.Frame frafondosRecon 
      Caption         =   "Fondos Recontados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4560
      Left            =   90
      TabIndex        =   22
      Top             =   2325
      Width           =   7890
      Begin Sicmact.FlexEdit fgOrdenPago 
         Height          =   1050
         Left            =   4110
         TabIndex        =   7
         Top             =   405
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   1852
         Cols0           =   3
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Orden-Importe"
         EncabezadosAnchos=   "450-1500-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R"
         FormatosEdit    =   "0-0-2"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdBilletajes 
         Caption         =   "&Efectivo"
         Height          =   360
         Left            =   855
         TabIndex        =   6
         Top             =   735
         Width           =   1485
      End
      Begin Sicmact.FlexEdit fgResumen 
         Height          =   2415
         Left            =   705
         TabIndex        =   8
         Top             =   1980
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   4260
         Cols0           =   3
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Concepto-Importe"
         EncabezadosAnchos=   "450-4000-1800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R"
         FormatosEdit    =   "0-0-2"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   7860
         X2              =   0
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   7890
         X2              =   30
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "I. Billetaje - Efectivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   150
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotalOrden 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6255
         TabIndex        =   27
         Top             =   1515
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ordenes :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4935
         TabIndex        =   26
         Top             =   1545
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "II. Ordenes de Pago Pendientes de Cobro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   4110
         TabIndex        =   25
         Top             =   195
         Width           =   3495
      End
      Begin VB.Label lblTotalEfec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2580
         TabIndex        =   24
         Top             =   1515
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Efectivo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1350
         TabIndex        =   23
         Top             =   1545
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   1215
         Top             =   1500
         Width           =   2670
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   4830
         Top             =   1500
         Width           =   2730
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
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
      Height          =   1035
      Left            =   75
      TabIndex        =   16
      Top             =   0
      Width           =   7890
      Begin VB.ListBox lstCajaPend 
         Height          =   480
         Left            =   6660
         TabIndex        =   0
         Top             =   315
         Width           =   915
      End
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   195
         TabIndex        =   1
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPersCodResp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   33
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblRespCajaChica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1770
         TabIndex        =   32
         Top             =   630
         Width           =   4380
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5625
         TabIndex        =   29
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1260
         TabIndex        =   17
         Top             =   270
         Width           =   4320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proc. Rend  "
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6675
         TabIndex        =   21
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   60
      TabIndex        =   12
      Top             =   960
      Width           =   7890
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   2880
         TabIndex        =   30
         Top             =   885
         Width           =   4635
         Begin MSComCtl2.DTPicker txtHoraIni 
            Height          =   330
            Left            =   930
            TabIndex        =   4
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   135528450
            CurrentDate     =   37097
         End
         Begin MSComCtl2.DTPicker txtHoraFin 
            Height          =   330
            Left            =   3210
            TabIndex        =   5
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   135528450
            CurrentDate     =   37097
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hora Final :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2325
            TabIndex        =   34
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hora Inicio :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   45
            TabIndex        =   31
            Top             =   105
            Width           =   840
         End
      End
      Begin Sicmact.TxtBuscar txtBuscarResp 
         Height          =   345
         Left            =   1395
         TabIndex        =   3
         Top             =   548
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
         ForeColor       =   0
      End
      Begin Sicmact.TxtBuscar txtBuscarAge 
         Height          =   330
         Left            =   1395
         TabIndex        =   2
         Top             =   202
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   0
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dd/mm/yyyy"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1395
         TabIndex        =   20
         Top             =   923
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   19
         Top             =   975
         Width           =   540
      End
      Begin VB.Label lblNombRespArqueo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000006&
         Height          =   315
         Left            =   2940
         TabIndex        =   18
         Top             =   563
         Width           =   4770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Resp. Arqueo :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2955
         TabIndex        =   14
         Top             =   210
         Width           =   4770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCajaChicaArqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oArendir As NARendir
Dim oCajaCH As nCajaChica
Dim oOpe As DOperacion

Dim lsCtaCajaChica As String
Dim lsCtaDifPos As String
Dim lsCtaDifNeg As String

Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdAceptar_Click()
Dim lsCtaHaber As String
Dim lsCtaDebe As String
Dim lnMonto As Currency
Dim lsFechaFin As String
Dim lsMovNro As String
Dim oCont As NContFunciones
Dim rsCH As ADODB.Recordset

Set oCont = New NContFunciones
Set rsCH = New ADODB.Recordset
    
If Valida = False Then Exit Sub

lnMonto = Me.fgResumen.TextMatrix(7, 2)
lsCtaCajaChica = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtBuscarAreaCH, ObjCMACAgenciaArea)
lsCtaDifPos = oOpe.EmiteOpeCta(gsOpeCod, "H")
lsCtaDifNeg = oOpe.EmiteOpeCta(gsOpeCod, "H", 1, gsCodAge, ObjCMACAgencias)

If MsgBox("Desea Realizar el Arqueo de caja chica", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lsFechaFin = Me.txtHoraFin
    MsgBox "Deberá realizar la Rendición de Caja Chica Actual. Para Poder Proseguir", vbExclamation, "Aviso"
    frmCajaChicaRendicion.Inicio gCHTipoProcArqueo, txtBuscarAreaCH
    If frmCajaChicaRendicion.OK = False Then
        Exit Sub
    End If
    If lnMonto > 0 Then
        lsCtaDebe = lsCtaDifPos
        lsCtaHaber = lsCtaCajaChica
    Else
        lsCtaDebe = lsCtaCajaChica
        lsCtaHaber = lsCtaDifNeg
    End If
    Set rsCH = CreaRsCH
    
    oCajaCH.GrabaArqueoCajaChica gsFormatoFecha, lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, _
                lsCtaHaber, rsBill, rsMon, fgResumen.GetRsNew, fgOrdenPago.GetRsNew, _
                rsCH, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), txtBuscarAge, _
                txtBuscarResp, txtHoraIni, lsFechaFin, lnMonto
    
    Dim oContImp As NContImprimir
    Dim lsTexto As String
    Set oContImp = New NContImprimir
    If Not rsBill Is Nothing Then
        rsBill.MoveFirst
    End If
    If Not rsMon Is Nothing Then
        rsMon.MoveFirst
    End If
    lsTexto = ""
    lsTexto = oContImp.ImprimeArqueo(gsNomCmac, gdFecSis, lsMovNro, gsOpeCod, rsBill, rsMon, _
                fgResumen.GetRsNew, fgOrdenPago.GetRsNew, txtBuscarAreaCH, lblCajaChicaDesc, _
                lblPersCodResp, lblRespCajaChica, txtBuscarAge, lblAgeDesc, txtBuscarResp, _
                lblNombRespArqueo, txtHoraIni, lsFechaFin, _
                gnLinPage, gnColPage)
    
    EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
    ImprimeAsientoContable lsMovNro
    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo Arqueo de Caja Chica de la  Agencia/Area : " & lblCajaChicaDesc & " |Responsable : " & lblNombRespArqueo
            Set objPista = Nothing
            '*******
    If MsgBox("Desea Realizar Arqueo de Otra Caja Chica??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Limpiar
        lblNroProcCH = ""
        lblCajaChicaDesc = ""
        txtBuscarAreaCH = ""
        txtBuscarAge = gsCodAge
        lblAgeDesc = gsNomAge
        lblPersCodResp = ""
        lblRespCajaChica = ""
        lstCajaPend.Clear
    Else
        Unload Me
    End If
End If
End Sub
Public Function Valida() As Boolean
Valida = True
If Len(Trim(txtBuscarAreaCH)) = 0 Then
    MsgBox "Caja chica no seleccionada", vbInformation, "Aviso"
    txtBuscarAreaCH.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtBuscarAge)) = 0 Then
    MsgBox "Agencia no ha sido seleccionada o no válida", vbInformation, "Aviso"
    If txtBuscarAge.Enabled Then txtBuscarAge.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtBuscarResp)) = 0 Then
    MsgBox "Responsable de Arqueo no ha sido ingresado", vbInformation, "Aviso"
    txtBuscarResp.SetFocus
    Valida = False
    Exit Function
End If
If Val(lblTotalEfec) = 0 Then
    If MsgBox("Efectivo no ha sido Ingresado. Desea Continuar??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdBilletajes.SetFocus
        Valida = False
        Exit Function
    End If
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no Ingresado", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Valida = False
    Exit Function
End If
If CDate(Format(lblFecha & " " & CDate(txtHoraIni), "dd/mm/yyyy hh:mm:ss AMPM")) >= CDate(Format(lblFecha & " " & CDate(txtHoraFin), "dd/mm/yyyy hh:mm:ss AMPM")) Then
    MsgBox "Hora de Final no puede ser menor que Hora Final", vbInformation, "aviso"
    Valida = False
    txtHoraFin.SetFocus
    Exit Function
End If

End Function

Private Sub cmdBilletajes_Click()
    Set rsBill = New ADODB.Recordset
    Set rsMon = New ADODB.Recordset
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, 0, Mid(gsOpeCod, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBill = frmCajaGenEfectivo.rsBilletes
        Set rsMon = frmCajaGenEfectivo.rsMonedas
        lblTotalEfec = frmCajaGenEfectivo.lblTotal
        fgResumen.TextMatrix(1, 2) = Format(CCur(lblTotalEfec) + CCur(lblTotalOrden), "#,#0.00")
        CalculaTotal
        txtMovDesc.SetFocus
    Else
        Set rsBill = Nothing
        Set rsMon = Nothing
    End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oAreas As DActualizaDatosArea

Set oOpe = New DOperacion
Set oArendir = New NARendir
Set oCajaCH = New nCajaChica
Set oAreas = New DActualizaDatosArea


        
CentraForm Me
lblFecha = gdFecSis
txtHoraIni = Time
txtHoraFin = Time

CargaFlex
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
txtBuscarAge.rs = oAreas.GetAgencias
txtBuscarAge.Text = gsCodAge
lblAgeDesc = gsNomAge
txtBuscarAge.Enabled = False
End Sub
Function GetCajasChicas(Optional lbActual As Boolean = True) As String
Dim i As Integer
Dim lsCajas As String
If lstCajaPend.ListCount > 0 Then
    For i = 0 To lstCajaPend.ListCount - 1
        lsCajas = lsCajas + lstCajaPend.List(i) + ","
    Next
End If
If lbActual Then
    lsCajas = lsCajas + lblNroProcCH
Else
    If lsCajas <> "" Then
        lsCajas = Mid(lsCajas, 1, Len(lsCajas) - 1)
    End If
End If
GetCajasChicas = lsCajas
End Function
Sub CargaFlex()
Dim i As Integer
fgResumen.Clear
fgResumen.FormaCabecera
fgResumen.Rows = 2
For i = 1 To 9
    fgResumen.AdicionaFila
    fgResumen.TextMatrix(i, 2) = "0.00"
    Select Case i
        Case 1
            fgResumen.TextMatrix(i, 1) = "TOTAL EFECTIVO"
        Case 2
            fgResumen.TextMatrix(i, 1) = "TOTAL DOCUMENTOS"
        Case 3
            fgResumen.TextMatrix(i, 1) = "DOCS. REPORTADOS A CONTABILIDAD PARA REEMBOLSO"
        Case 4
            fgResumen.TextMatrix(i, 1) = "REG. A RENDIR CUENTA"
            fgResumen.BackColorRow &HFF8080, True
        Case 5
            fgResumen.TextMatrix(i, 1) = "TOTAL DOC.SUSTENTADOS - A RENDIR CTA"
        Case 6
            fgResumen.TextMatrix(i, 1) = "TOTAL DOC. NO SUSTENTADOS - A RENDIR CTA"
        Case 7
            fgResumen.TextMatrix(i, 1) = "TOTAL FONDO FIJO SEGUN ARQUEO"
            fgResumen.BackColorRow &HE0E0E0, True
        Case 8
            fgResumen.TextMatrix(i, 1) = "FONDO ASIGNADO"
        Case 9
            fgResumen.TextMatrix(i, 1) = "DIFERENCIA"
            fgResumen.BackColorRow &HE0E0E0, True
    End Select
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oCajaCH = Nothing
Unload frmCajaGenEfectivo
Set frmCajaGenEfectivo = Nothing
End Sub

Private Sub txtBuscarAreaCH_EmiteDatos()
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
CargaDatos
If txtBuscarResp.Visible Then
    txtBuscarResp.SetFocus
End If
End Sub
Private Sub txtBuscarResp_EmiteDatos()
lblNombRespArqueo = txtBuscarResp.psDescripcion
If txtBuscarResp.Text <> "" Then
    usu.DatosPers txtBuscarResp.Text
    lblNombRespArqueo = ""
    If usu.PersCod = "" Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
    Else
        lblNombRespArqueo = txtBuscarResp.psDescripcion
        txtHoraIni.SetFocus
    End If
End If

End Sub
Public Sub CargaDatos()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Limpiar
Set rs = oCajaCH.GetRSDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
If Not rs.EOF And Not rs.BOF Then
    lblPersCodResp = rs!cPersCod
    lblRespCajaChica = PstaNombre(rs!cPersNombre)
End If
rs.Close
Set rs = Nothing
Set rs = oCajaCH.GetCHArqueoRendidos(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
lstCajaPend.Clear
If Not rs.EOF And Not rs.BOF Then
    lstCajaPend.AddItem rs!nProcNro
End If
fgResumen.TextMatrix(8, 2) = Format(oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoAsig), "#,#0.00")
fgResumen.TextMatrix(2, 2) = Format(oCajaCH.GetTotalDocArqueos(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Trim(Str(lblNroProcCH))), "#,#0.00")
fgResumen.TextMatrix(4, 2) = Format(oCajaCH.GetTotalArendirArqueos(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), GetCajasChicas), "#,#0.00")
fgResumen.TextMatrix(5, 2) = Format(oCajaCH.GetTotalArendirArqueosSustentados(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), GetCajasChicas), "#,#0.00")
fgResumen.TextMatrix(6, 2) = Format(oCajaCH.GetTotalArendirArqueosNOSustentados(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), GetCajasChicas), "#,#0.00")

If GetCajasChicas(False) <> "" Then
    fgResumen.TextMatrix(3, 2) = Format(oCajaCH.GetTotalDocArqueos(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), GetCajasChicas(False)), "#,#0.00")
End If
Set rs = oCajaCH.GetOrdenesNoCobradas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), GetCajasChicas)
fgOrdenPago.Clear
fgOrdenPago.FormaCabecera
fgOrdenPago.Rows = 2
lblTotalOrden = "0.00"
If Not rs.EOF And Not rs.BOF Then
    Set fgOrdenPago.Recordset = rs
    lblTotalOrden = Format(fgOrdenPago.SumaRow(2), "#,#0.00")
End If
fgResumen.TextMatrix(1, 2) = Format(CCur(lblTotalEfec) + CCur(lblTotalOrden), "#,#0.00")
CalculaTotal
End Sub
Sub CalculaTotal()
Dim lnTotalEfec As Currency
Dim lnTotalDoc As Currency
Dim lnDocContab As Currency
Dim lnARendir As Currency
Dim lnTotalFondo As Currency
Dim lnFondoAsig As Currency
Dim lnDiferencia As Currency
Dim lnSustentado As Currency
Dim lnNSustentado As Currency

lnTotalEfec = CCur(fgResumen.TextMatrix(1, 2))
lnTotalDoc = CCur(fgResumen.TextMatrix(2, 2))
lnDocContab = CCur(fgResumen.TextMatrix(3, 2))
lnARendir = CCur(fgResumen.TextMatrix(4, 2))
lnFondoAsig = CCur(fgResumen.TextMatrix(8, 2))
lnSustentado = CCur(fgResumen.TextMatrix(5, 2))
lnNSustentado = CCur(fgResumen.TextMatrix(6, 2))

'lnTotalFondo = lnTotalEfec + lnTotalDoc + lnDocContab + lnARendir - lnSustentado
lnTotalFondo = lnTotalEfec + lnTotalDoc + lnDocContab + lnSustentado + lnNSustentado

lnDiferencia = lnFondoAsig - lnTotalFondo
'lnDiferencia = lnFondoAsig - lnTotalFondo
fgResumen.TextMatrix(7, 2) = Format(lnTotalFondo, "#,#0.00")
fgResumen.TextMatrix(9, 2) = Format(lnDiferencia, "#,#0.00")

End Sub
Sub Limpiar()
Me.fgOrdenPago.Clear
Me.fgOrdenPago.FormaCabecera
Me.fgOrdenPago.Rows = 2
CargaFlex
Unload frmCajaGenEfectivo
Set frmCajaGenEfectivo = Nothing
lblTotalEfec = "0.00"
lblTotalOrden = "0.00"
txtBuscarResp = ""
lblNombRespArqueo = ""
txtHoraIni = Time
txtHoraFin = Time
End Sub
Function CreaRsCH() As ADODB.Recordset
Dim i As Integer
Dim lsCajas As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Fields.Append "NroCH", adVarChar, 12
rs.Open
If lstCajaPend.ListCount > 0 Then
    For i = 0 To lstCajaPend.ListCount - 1
        rs.AddNew
        rs!NroCH = lstCajaPend.List(i)
        rs.Update
    Next
End If
rs.AddNew
rs!NroCH = lblNroProcCH
rs.Update
rs.MoveFirst
Set CreaRsCH = rs
End Function
Private Sub txtHoraIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHoraFin.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdAceptar.SetFocus
End If
End Sub
