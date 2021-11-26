VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmArqueoTarjDebBoveda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueo de Stock de Tarjetas de Débito - Bóveda"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   Icon            =   "frmArqueoTarjDebBoveda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegArqueo 
      Caption         =   "Registrar Arqueo"
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
      Left            =   10080
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   12000
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Stock de Tarjeta en Bóveda"
      TabPicture(0)   =   "frmArqueoTarjDebBoveda.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUsuArqField"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUsuArqValue"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUsuSupField"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblUsuSupValue"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFechaField"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFechaValue"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "feStockTarjBove"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraGlosa"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbClase"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Arqueo de Tarjetas Habilitadas"
      TabPicture(1)   =   "frmArqueoTarjDebBoveda.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feTarjHabilitadas"
      Tab(1).Control(1)=   "cmdGenerar"
      Tab(1).Control(2)=   "fraPeriodo"
      Tab(1).Control(3)=   "fraUsuReceptor"
      Tab(1).ControlCount=   4
      Begin VB.ComboBox cmbClase 
         Height          =   315
         ItemData        =   "frmArqueoTarjDebBoveda.frx":0342
         Left            =   10200
         List            =   "frmArqueoTarjDebBoveda.frx":034C
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   600
         Width           =   2775
      End
      Begin SICMACT.FlexEdit feTarjHabilitadas 
         Height          =   4215
         Left            =   -74640
         TabIndex        =   20
         Top             =   1560
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7435
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Tarjeta-Usuario Receptor-Ok"
         EncabezadosAnchos=   "300-1800-1800-1200"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   375
         Left            =   -65520
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Frame fraPeriodo 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69000
         TabIndex        =   16
         Top             =   600
         Width           =   3255
         Begin MSComCtl2.DTPicker dtpPerIni 
            Height          =   300
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   141492225
            CurrentDate     =   39743
         End
         Begin MSComCtl2.DTPicker dtpPerFin 
            Height          =   300
            Left            =   1680
            TabIndex        =   22
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   141492225
            CurrentDate     =   39743
         End
         Begin VB.Label Label5 
            Caption         =   "--"
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
            Left            =   1480
            TabIndex        =   18
            Top             =   400
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "-"
            Height          =   15
            Left            =   1560
            TabIndex        =   17
            Top             =   480
            Width           =   135
         End
      End
      Begin VB.Frame fraUsuReceptor 
         Caption         =   "Usuario Receptor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   12
         Top             =   600
         Width           =   5535
         Begin VB.CheckBox chkTodos 
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
            Left            =   1690
            TabIndex        =   13
            Top             =   0
            Width           =   930
         End
         Begin SICMACT.TxtBuscar txtReceptorCod 
            Height          =   285
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.Label lblUsuReceptorValue 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1680
            TabIndex        =   15
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   10
         Top             =   5160
         Width           =   12735
         Begin VB.TextBox txtGlosa 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   405
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   12495
         End
      End
      Begin SICMACT.FlexEdit feStockTarjBove 
         Height          =   4215
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7435
         Cols0           =   11
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Precinto-Cod. Lote-Cod. Sub Lote-Tarj. Inicial-Tarj. Final-Cant. Stock-Cant. Física-Detalle-Aux-Estado"
         EncabezadosAnchos=   "300-1200-1250-1300-1600-1600-1200-1200-1400-0-1500"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-7-8-9-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-1"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-3-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Clase:"
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
         Left            =   9360
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblFechaValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7080
         TabIndex        =   8
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label lblFechaField 
         Caption         =   "Fecha:"
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
         Left            =   6360
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblUsuSupValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5040
         TabIndex        =   6
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label lblUsuSupField 
         Caption         =   "Usuario Supervisa:"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblUsuArqValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2040
         TabIndex        =   4
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label lblUsuArqField 
         Caption         =   "Usuario Arqueado:"
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
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmArqueoTarjDebBoveda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmArqueoTarjDebBoveda
'** Descripción : Formulario para realizar el Arqueo de Stock de Tarjetas de Debito  - Boveda
'** Creación : PASI, 20151221
'** Referencia : TI-ERS069-2015
'***************************************************************************
Option Explicit
Dim bResultadoVisto As Boolean
Dim oVisto As frmVistoElectronico
Dim cUsuVisto As String
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim bConforme As Boolean 'GIPO 30/09/16 ERS051-2016
Dim textoEstado As String 'GIPO 04/10/16 ERS051-2016

Dim nMatDetFaltante() As TDetFaltante
Private Type TDetFaltante
    cPrecinto As String
    cCodLote As String
    cCodSubLote As String
    cNumTarjeta As String
    nFaltante As Integer
End Type
Public Sub Inicia()
    Set oVisto = New frmVistoElectronico
    bResultadoVisto = oVisto.Inicio(15)
    If Not bResultadoVisto Then
        Exit Sub
    End If
    cUsuVisto = oVisto.ObtieneUsuarioVisto
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    If oCaja.ObtieneExisteArqueoBoveda(gdFecSis, UCase(gsCodUser)) Then
        MsgBox "El arqueo de este día ya ha sido realizado.", vbInformation, "Mensaje"
        Exit Sub
    End If
    Me.Show 1
End Sub
Private Sub chkTodos_Click()
    txtReceptorCod.Enabled = IIf(chkTodos.value = 1, False, True)
    txtReceptorCod.Text = IIf(chkTodos.value = 1, "", txtReceptorCod.Text)
    txtReceptorCod.psCodigoPersona = IIf(chkTodos.value = 1, "", txtReceptorCod.psCodigoPersona)
    lblUsuReceptorValue.Caption = IIf(chkTodos.value = 1, "", lblUsuReceptorValue.Caption)
    LimpiaFlex feTarjHabilitadas
End Sub

'GIPO ERS051-2016 03/10/2016
Private Sub cmbClase_Click()
    Me.txtGlosa.Text = textoEstado & "/" & Me.cmbClase.Text
End Sub
'END GIPO

Private Sub cmdCancelar_Click()
    LimpiaFlex feTarjHabilitadas
    LimpiaFlex feStockTarjBove
    chkTodos.value = 0
    dtpPerIni.value = CDate(gdFecSis)
    dtpPerFin.value = CDate(gdFecSis)
    txtGlosa.Text = ""
    CargaDatos
End Sub

Private Sub cmdGenerar_Click()
LimpiaFlex feTarjHabilitadas 'GIPO
Dim rs As ADODB.Recordset
    If Not ValidaTabHabilita Then Exit Sub
    Set rs = oCaja.ObtieneTarjetasHabxArqueo(lblUsuArqValue.Caption, IIf(chkTodos.value = 1, "%", txtReceptorCod.Text), CDate(dtpPerIni.value), CDate(dtpPerFin.value))
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos para mostrar. Verifique", vbInformation, "Mensaje"
        Exit Sub
    End If
    Do While Not rs.EOF
        feTarjHabilitadas.AdicionaFila
        feTarjHabilitadas.TextMatrix(feTarjHabilitadas.row, 1) = rs!cNumTarjeta
        feTarjHabilitadas.TextMatrix(feTarjHabilitadas.row, 2) = rs!cUsuReceptor
        rs.MoveNext
    Loop
End Sub
Private Function ValidaTabHabilita() As Boolean
    ValidaTabHabilita = False
    If chkTodos.value = 0 Then
        If Len(Trim(txtReceptorCod.Text)) = 0 Then
            MsgBox "Asegurese de haber ingresado el Usuario Receptor.", vbInformation, "Mensaje"
            txtReceptorCod.SetFocus
            Exit Function
        End If
        If Len(Trim(lblUsuReceptorValue.Caption)) = 0 Then
            MsgBox "Al parecer los datos del Usuario Receptor no se han cargado correctamente. Verifique.", vbInformation, "Mensaje"
            txtReceptorCod.SetFocus
            Exit Function
        End If
    End If
    If Not IsDate(dtpPerIni.value) Then
        MsgBox "La Fecha de Inicio no es válida. Verifique.", vbInformation, "Mensaje"
        dtpPerIni.SetFocus
        Exit Function
    End If
    If Not IsDate(dtpPerFin.value) Then
        MsgBox "La Fecha de Fin no es válida. Verifique.", vbInformation, "Mensaje"
        dtpPerFin.SetFocus
        Exit Function
    End If
    If CDate(dtpPerIni.value) > CDate(dtpPerFin.value) Then
        MsgBox "La Fecha Inicio no puede ser superior a la Fecha Fin. Verifique.", vbInformation, "Mensaje"
        dtpPerIni.SetFocus
        Exit Function
    End If
    ValidaTabHabilita = True
End Function
Private Sub cmdRegArqueo_Click()
On Error GoTo ErrorRegistra
Dim oDMov As DMov
Dim i As Integer
Dim X As Integer
Dim lsMovNro As String
Dim nIdArqueo As Integer
Dim nIdArqueoBov As Integer
Dim bTrans As Boolean

Me.cmdRegArqueo.Enabled = False 'GIPO

Set oDMov = New DMov
If Not ValidaArqueo Then
    Me.cmdRegArqueo.Enabled = True
    Exit Sub
End If
If feStockTarjBove.TextMatrix(1, 1) = "" Then
    If MsgBox("No existen datos en el Stock de Tarjetas en Bóveda. Está seguro de continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Me.cmdRegArqueo.Enabled = True
        Exit Sub
    End If
End If
If feTarjHabilitadas.TextMatrix(1, 1) = "" Then
    If MsgBox("No existen Datos para el Arqueo de Tarjetas Habilitadas. Está Seguro de continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Me.cmdRegArqueo.Enabled = True
        Exit Sub
    End If
End If
If feStockTarjBove.TextMatrix(1, 1) <> "" Then
    For i = 1 To feStockTarjBove.Rows - 1
        If Trim(feStockTarjBove.TextMatrix(i, 8)) = "" Then
            MsgBox "En el Stock de Tarjeta en Bóveda no se ha ingresado el detalle del registro (" & i & "). Verifique.", vbInformation, "Mensaje"
            Me.cmdRegArqueo.Enabled = True
            Exit Sub
        End If
    Next
End If
    If MsgBox("¿Está seguro de realizar el Arqueo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Me.cmdRegArqueo.Enabled = True
        Exit Sub
    End If
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, UCase(gsCodUser))
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    oCaja.dBeginTrans
    bTrans = True
    'GIPO ERS051-2016
    nIdArqueo = oCaja.RegistrarArqueoTarjDebito(lblUsuArqValue.Caption, lblUsuSupValue.Caption, gdFecSis & " " & GetHoraServer, lsMovNro, Trim(Replace(Replace((txtGlosa.Text), Chr(10), ""), Chr(13), "")))
    For i = 1 To feStockTarjBove.Rows - 1
        nIdArqueoBov = oCaja.RegistrarArqueoTarjDebitoEnBoveda(nIdArqueo, feStockTarjBove.TextMatrix(i, 1), feStockTarjBove.TextMatrix(i, 2), feStockTarjBove.TextMatrix(i, 3), feStockTarjBove.TextMatrix(i, 4), feStockTarjBove.TextMatrix(i, 5), feStockTarjBove.TextMatrix(i, 6), feStockTarjBove.TextMatrix(i, 7))
        For X = 1 To UBound(nMatDetFaltante)
            If nMatDetFaltante(X).cPrecinto = feStockTarjBove.TextMatrix(i, 1) And _
                nMatDetFaltante(X).cCodLote = feStockTarjBove.TextMatrix(i, 2) And _
                nMatDetFaltante(X).cCodSubLote = feStockTarjBove.TextMatrix(i, 3) Then
                oCaja.RegistrarArqueoTarjDebitoEnBovedaDet nIdArqueoBov, nMatDetFaltante(X).cNumTarjeta, nMatDetFaltante(X).nFaltante
            End If
        Next
    Next
    oCaja.RegistrarArqueoTarjDebitoHabilita nIdArqueo, chkTodos.value, CDate(dtpPerIni.value), CDate(dtpPerFin.value)
    For i = 1 To feTarjHabilitadas.Rows - 1
        oCaja.RegistrarArqueoTarjDebitoHabilitaDet nIdArqueo, feTarjHabilitadas.TextMatrix(i, 1), feTarjHabilitadas.TextMatrix(i, 2), IIf(feTarjHabilitadas.TextMatrix(i, 3) = ".", 1, 0)
    Next
    oCaja.dCommitTrans
    'MARG ERS052-2017----
       oVisto.RegistraVistoElectronico 0, , lblUsuArqValue.Caption
    'END MARG- -------------
       
    MsgBox "El Arqueo ha sido realizado correctamente.", vbInformation, "Aviso"
    'imprimirPDF GIPO
    Call generarPDFarqueo(nIdArqueo) 'GIPO ERS051-2016
    bTrans = False
    Unload Me
Exit Sub
ErrorRegistra:
    If bTrans Then
        oCaja.dRollbackTrans
        Set oCaja = Nothing
    End If
    MsgBox err.Number & " - " & err.Description, vbInformation, "Error"
End Sub
Private Function ValidaArqueo() As Boolean
    ValidaArqueo = False
    If Len(lblUsuArqValue.Caption) = 0 Then
        MsgBox "No se encuentra el Usuario Arqueado. Verifique", vbInformation, "Mensaje"
        Exit Function
    End If
    If Len(lblUsuSupValue.Caption) = 0 Then
        MsgBox "No se encuentra el Usuario Supervisor. Verifique", vbInformation, "Mensaje"
        Exit Function
    End If
    If Not IsDate(lblFechaValue.Caption) Then
        MsgBox "La Fecha de Arqueo no es válida. Verifique", vbInformation, "Mensaje"
        Exit Function
    End If
    If Len(txtGlosa.Text) = 0 Then
        MsgBox "No se ha ingresado la Glosa. Verifique", vbInformation, "Mensaje"
        Exit Function
    End If
    If chkTodos.value = 0 Then
        If Len(Trim(txtReceptorCod.Text)) = 0 Then
            MsgBox "Para el Arqueo de Tarjetas Habilitadas asegurese de haber ingresado el Usuario Receptor.", vbInformation, "Mensaje"
            Exit Function
            
        End If
        If Len(Trim(lblUsuReceptorValue.Caption)) = 0 Then
            MsgBox "Para el Arqueo de Tarjetas Habilitadas asegurese que los datos del Usuario Receptor se han cargado correctamente. Verifique.", vbInformation, "Mensaje"
            Exit Function
        End If
    End If
    ValidaArqueo = True
End Function
'Created By GIPO 04-11-2016
Public Sub generarPDFarqueo(ByVal IdArqueo As Integer)
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    'idArqueo = 318
    Set rs = oCaja.obtenerDatosActaTarjArqueadas(IdArqueo)
    Set rs1 = oCaja.obtenerDetalleTarjArqueadasBoveda(IdArqueo)
    Set Rs2 = oCaja.obtenerDetalleTarjArqueadasHabilitadasBoveda(IdArqueo)
    
    Dim oDoc As cPDF
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\Spooler\ACTA_ARQUEO_BOVEDA_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    
    oDoc.WImage 75, 40, 35, 105, "Logo"
    oDoc.WTextBox 63, 40, 15, 500, "ACTA DE ARQUEO DE TARJETAS VISA - BÓVEDA", "F2", 12, hCenter
    oDoc.WTextBox 90, 40, 732, 510, "En la ciudad de " & rs!Ciudad & " , el día " & rs!FechaCompleta & ", " & _
                                    "a horas " & rs!Hora & " se suscribe en actas que en Bóveda de la " & rs!Agencia & "," & _
                                    "con la presencia del (la) Sr(a). " & rs!NombrePersonaArqueada & " (" & rs!CargoUserArqueado & ") " & _
                                    "y el (la) Sr(a). " & rs!NombreArqueador & " (" & rs!CargoUserSuperviza & "), se procedió a realizar " & _
                                    "el arqueo de tarjetas débito Visa en la Agencia mencionada.", "F1", 9, hjustify, vTop, vbBlack, 0, vbBlack, False, 10
                                    
    oDoc.WTextBox 145, 50, 20, 510, "DETALLE DE TARJETAS ARQUEADAS EN BÓVEDA", "F2", 9, hjustify
    
    oDoc.WTextBox 165, 50, 20, 80, "Precinto", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox 165, 130, 20, 80, "Cod. Lote", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox 165, 210, 20, 80, "Cod. SubLote", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox 165, 290, 20, 100, "Tarj. Inicial", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox 165, 390, 20, 100, "Tarj. Final", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    Dim i As Integer
    Dim h1 As Integer
    h1 = 0
    Do While Not rs1.EOF
        h1 = h1 + 20
        oDoc.WTextBox 165 + h1, 50, 20, 80, rs1!cCodPrecinto, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox 165 + h1, 130, 20, 80, rs1!cCodLote, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox 165 + h1, 210, 20, 80, rs1!cCodSubLote, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox 165 + h1, 290, 20, 100, rs1!cRangTarjDel, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox 165 + h1, 390, 20, 100, rs1!cRangTarjAl, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        rs1.MoveNext
    Loop
    h1 = 165 + h1 + 10 'espacio después del cuadro
    
    'GIPO MEJORA SEGÚN INC1709060018 30-09-2017
    h1 = h1 + 40
     If h1 >= 700 Then
      'imprimir pie de página antes de crear la siguiente
        oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
        oDoc.NewPage A4_Vertical
        'h1 = h1 - 650
        h1 = 50
    End If
    
    '************************DETALLE DE TARJETAS HABILITADAS**************************
    oDoc.WTextBox h1 + 30, 50, 20, 510, "DETALLE DE TARJETAS HABILITADAS", "F2", 9, hjustify
    
    oDoc.WTextBox h1 + 50, 50, 20, 100, "Número Tarjeta", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox h1 + 50, 150, 20, 80, "Usuario Receptor", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox h1 + 50, 230, 20, 80, "Estado", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    
    h1 = h1 + 50
    Do While Not Rs2.EOF
        h1 = h1 + 20
        If h1 >= 745 Then
            'imprimir pie de página antes de crear la siguiente
            oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
            oDoc.NewPage A4_Vertical
            h1 = h1 - 705
        End If
        oDoc.WTextBox h1, 50, 20, 100, Rs2!cNumTarjeta, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1 'GIPO 20181205 corrección del Dígito de más
        oDoc.WTextBox h1, 150, 20, 80, Rs2!cUserArqueado, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox h1, 230, 20, 80, Rs2!Estado, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        Rs2.MoveNext
    Loop
    
    '************************PÁRRAFO FINAL**************************
    h1 = h1 + 40
     If h1 >= 700 Then
      'imprimir pie de página antes de crear la siguiente
        oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
        oDoc.NewPage A4_Vertical
        h1 = h1 - 650
    End If
    
    Dim parrafoFinal As String
    parrafoFinal = "Siendo las " & rs!Hora & " del día " & rs!FechaCompleta & ", se dió por concluido " & _
                   "el arqueo de tarjetas de débito VISA; por lo cual firman en señal de conformidad."
    oDoc.WTextBox h1, 50, 20, 510, parrafoFinal, "F1", 9, hjustify
    
    'GIPO MEJORA SEGÚN INC1709060018 30-09-2017
    h1 = h1 + 40
     If h1 >= 700 Then
      'imprimir pie de página antes de crear la siguiente
        oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
        oDoc.NewPage A4_Vertical
        'h1 = h1 - 650
        h1 = 50
    End If

    
    h1 = h1 + 70
    oDoc.WTextBox h1, 50, 20, 250, "_________________________________", "F1", 9, hCenter
    oDoc.WTextBox h1 + 15, 50, 20, 250, rs!NombreArqueador, "F2", 9, hCenter
    oDoc.WTextBox h1 + 25, 50, 20, 250, rs!CargoUserSuperviza, "F2", 9, hCenter
    oDoc.WTextBox h1 + 35, 50, 20, 250, "Responsable Arqueo", "F2", 9, hCenter
    
    oDoc.WTextBox h1, 300, 20, 250, "_________________________________", "F1", 9, hCenter
    oDoc.WTextBox h1 + 15, 300, 20, 250, rs!NombrePersonaArqueada, "F2", 9, hCenter
    oDoc.WTextBox h1 + 25, 300, 20, 250, rs!CargoUserArqueado, "F2", 9, hCenter
    oDoc.WTextBox h1 + 35, 300, 20, 250, "Usuario Arqueado", "F2", 9, hCenter
    
    
    '************************PIE DE PÁGINA**************************
    Dim sPiePagina As String
    sPiePagina = printFooter
    oDoc.WTextBox 767, 45, 15, 510, sPiePagina, "F1", 7, hjustify
    
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Function printFooter()
Dim footer As String
footer = "Oficina Principal: Jr. Próspero No  791 -  Iquitos ;" & _
                 "Ag. Calle Arequipa: Ca Arequipa Nº 428;  Agencia Punchana Av. 28 de Julio 829 - Iquitos;" & _
                 "Ag. Belén: Av. Grau Nº 1260 - Iquitos; Ag. San Juan Bautista- Avda. Abelardo Quiñones Nº 2670- Iquitos;" & _
                 "Ag. Pucallpa: Jr. Ucayali  No  850 - 852 ; Ag. Huánuco: Jr. General Prado No   836;" & _
                 "Ag. Yurimaguas: Ca. Simón Bolívar Nº 113; Ag. Tingo María: Av. Antonio Raymondi  Nº 246 ;" & _
                 "Ag. Tarapoto: Jr San Martín Nº 205 ;  Ag. Requena: Calle San Francisco Mz 28 Lt 07;" & _
                 "Ag. Cajamarca: Jr. Amalia Puga  Nº 417; Ag.  Aguaytía- Jr. Rio Negro Nº 259;" & _
                 "Ag. Cerro de Pasco; Plaza Carrión Nº 191; Ag. Minka: Ciudad Comercial Minka - Av. Argentina Nº 3093- Local 230 - Callao."
printFooter = footer
End Function

Private Sub feStockTarjBove_OnCellChange(pnRow As Long, pnCol As Long)
Call GetArqueoConforme 'GIPO
Dim i As Integer
    If pnCol = 7 Then
        If Trim(feStockTarjBove.TextMatrix(pnRow, 6)) = Trim(feStockTarjBove.TextMatrix(pnRow, 7)) Then
            feStockTarjBove.TextMatrix(pnRow, 8) = "OK"
            feStockTarjBove.TextMatrix(pnRow, 10) = "CONFORME"   'GIPO ERS051-2016
        Else
            feStockTarjBove.TextMatrix(pnRow, 8) = ""
            feStockTarjBove.TextMatrix(pnRow, 10) = "NO CONFORME" 'GIPO ERS051-2016
        End If
        
        For i = 1 To UBound(nMatDetFaltante)
            If nMatDetFaltante(i).cPrecinto = feStockTarjBove.TextMatrix(pnRow, 1) And _
                nMatDetFaltante(i).cCodLote = feStockTarjBove.TextMatrix(pnRow, 2) And _
                nMatDetFaltante(i).cCodSubLote = feStockTarjBove.TextMatrix(pnRow, 3) Then
                nMatDetFaltante(i).nFaltante = 0
            End If
        Next
        SendKeys "{Tab}", True
    End If
End Sub

'GIPO ERS051-2016
Private Sub GetArqueoConforme()
    Dim i As Integer
     For i = 1 To feStockTarjBove.Rows - 1
          If feStockTarjBove.TextMatrix(i, 6) <> feStockTarjBove.TextMatrix(i, 7) Then
            Me.txtGlosa.Text = "NO CONFORME/" & Me.cmbClase.Text
            textoEstado = "NO CONFORME"
            Exit Sub
          ElseIf feStockTarjBove.TextMatrix(i, 7) = "" Then
            Exit Sub
          End If
     Next
     Me.txtGlosa.Text = "CONFORME/" & Me.cmbClase.Text
     textoEstado = "CONFORME"
End Sub

Private Sub feStockTarjBove_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim row As Integer
    Dim vMat As Variant
    Dim vMatFalt As Variant
    Dim i As Integer
    Dim bHab As Boolean
    Dim X As Long
    Dim iMat As Integer
    
    Dim bConforme As Boolean  'GIPO ERS051-2016
    bConforme = False
    
    row = feStockTarjBove.row
    
    If Trim(feStockTarjBove.TextMatrix(row, 7)) = "" Then
        MsgBox "No se ha ingresado la Cantidad Fisica. Verifique.", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If Trim(feStockTarjBove.TextMatrix(row, 6)) = Trim(feStockTarjBove.TextMatrix(row, 7)) Then
        'MsgBox "El caso no requiere registrar faltante. Continue.", vbInformation, "Mensaje"
        psCodigo = "OK"
        bConforme = True
        'Exit Sub

    End If
    
    ReDim vMat(5, 0)
    For i = 1 To UBound(nMatDetFaltante)
        If nMatDetFaltante(i).cPrecinto = feStockTarjBove.TextMatrix(row, 1) And _
            nMatDetFaltante(i).cCodLote = feStockTarjBove.TextMatrix(row, 2) And _
            nMatDetFaltante(i).cCodSubLote = feStockTarjBove.TextMatrix(row, 3) Then
        
        iMat = UBound(vMat, 2) + 1
        ReDim Preserve vMat(5, 0 To iMat)
        vMat(1, iMat) = nMatDetFaltante(i).cPrecinto 'Precinto
        vMat(2, iMat) = nMatDetFaltante(i).cCodLote 'Codlote
        vMat(3, iMat) = nMatDetFaltante(i).cCodSubLote 'codSubLote
        vMat(4, iMat) = nMatDetFaltante(i).cNumTarjeta 'NumTarjeta
        vMat(5, iMat) = nMatDetFaltante(i).nFaltante 'nFaltante
        End If
    Next
    If UBound(vMat, 2) >= 1 Then
        bHab = False
        vMatFalt = frmArqueoTarjDebBovedaDetFal.Inicio(vMat, CLng(Trim(feStockTarjBove.TextMatrix(row, 6))) - CLng(Trim(feStockTarjBove.TextMatrix(row, 7))), bConforme)
        For i = 1 To UBound(nMatDetFaltante)
            For X = 1 To UBound(vMatFalt, 2)
                If vMatFalt(1, X) = nMatDetFaltante(i).cPrecinto And _
                    vMatFalt(2, X) = nMatDetFaltante(i).cCodLote And _
                    vMatFalt(3, X) = nMatDetFaltante(i).cCodSubLote And _
                    vMatFalt(4, X) = nMatDetFaltante(i).cNumTarjeta Then
                    nMatDetFaltante(i).nFaltante = vMatFalt(5, X)
                    If vMatFalt(5, X) = 1 Then
                        bHab = True
                    End If
                End If
            Next
        Next
        If bConforme = False Then
           psCodigo = IIf(bHab, "OK", "")
        End If
    End If
End Sub
Private Sub feStockTarjBove_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 7 Then
        If IsNumeric(feStockTarjBove.TextMatrix(feStockTarjBove.row, pnCol)) = False Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If CLng(feStockTarjBove.TextMatrix(feStockTarjBove.row, pnCol)) < 0 Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If CLng(feStockTarjBove.TextMatrix(feStockTarjBove.row, 6)) < CLng(feStockTarjBove.TextMatrix(feStockTarjBove.row, pnCol)) Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    CargaDatos
End Sub
Private Sub CargaDatos()
    Dim rs As ADODB.Recordset
    Dim rsNTarj As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rsNTarj = New ADODB.Recordset

    Me.cmbClase.ListIndex = 0 'GIPO ERS051-2016
    textoEstado = "" 'GIPO ERS051-2016
   
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
   
    Me.lblUsuArqValue.Caption = UCase(gsCodUser)
    Me.lblUsuSupValue.Caption = UCase(cUsuVisto)
    Me.lblFechaValue.Caption = gdFecSis
    dtpPerIni.value = CDate(gdFecSis)
    dtpPerFin.value = CDate(gdFecSis)
        
    ReDim Preserve nMatDetFaltante(0)
    
    Set rs = oCaja.ObtieneTarjetasDebitoEnBovedaxUser(gsCodUser)
    Do While Not rs.EOF
        feStockTarjBove.AdicionaFila
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 1) = rs!cCodPrecinto
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 2) = rs!cCodLote
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 3) = rs!cCodSubLote
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 4) = rs!cTarjIni
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 5) = rs!cTarjFin
        feStockTarjBove.TextMatrix(feStockTarjBove.row, 6) = rs!nCantTarj
        
        Set rsNTarj = oCaja.ObtieneTarjetasxArqueo(rs!cCodPrecinto, rs!cCodLote, rs!cCodSubLote, gsCodUser)
        Do While Not rsNTarj.EOF
            ReDim Preserve nMatDetFaltante(UBound(nMatDetFaltante) + 1)
            nMatDetFaltante(UBound(nMatDetFaltante)).cPrecinto = rs!cCodPrecinto
            nMatDetFaltante(UBound(nMatDetFaltante)).cCodLote = rs!cCodLote
            nMatDetFaltante(UBound(nMatDetFaltante)).cCodSubLote = rs!cCodSubLote
            nMatDetFaltante(UBound(nMatDetFaltante)).cNumTarjeta = rsNTarj!cNumTarjeta
            nMatDetFaltante(UBound(nMatDetFaltante)).nFaltante = 0
            rsNTarj.MoveNext
        Loop
        rs.MoveNext
    Loop
End Sub



Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdRegArqueo.SetFocus
    End If
End Sub
Private Sub txtReceptorCod_EmiteDatos()
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim R As New ADODB.Recordset
    Set ClsPersona = New COMDPersona.DCOMPersonas
    Set R = ClsPersona.BuscaCliente(txtReceptorCod.psCodigoPersona, BusquedaEmpleadoCodigo)
    If Not (R.EOF And R.BOF) Then
        If UCase(R!cUser) = lblUsuSupValue.Caption Then
            MsgBox "El Usuario Receptor es el mismo que el Supervisor. Verifique", vbInformation, "Mensaje"
            txtReceptorCod.Text = ""
            txtReceptorCod.psCodigoPersona = ""
            Exit Sub
        End If
        txtReceptorCod.Text = UCase(R!cUser)
        lblUsuReceptorValue.Caption = txtReceptorCod.psDescripcion
    End If
End Sub



