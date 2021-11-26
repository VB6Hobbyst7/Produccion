VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGenerarCarta 
   Caption         =   "Generar Carta de Circularización - Saldos"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerarCarta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   3495
      Left            =   9000
      TabIndex        =   33
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3495
      Left            =   9000
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
      Begin VB.CommandButton Command3 
         Caption         =   "I. Todos"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   8775
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   3135
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cFApertura"
            Caption         =   "F. Apertura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cTipoAho"
            Caption         =   "Producto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cAgeDescripcion"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cCtaCod"
            Caption         =   "Nro Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cPersNombre"
            Caption         =   "Titular"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nSaldoDisp"
            Caption         =   "Saldo Disp"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "sMoneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtDireccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   36
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   5640
         TabIndex        =   32
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtDeImporte 
         Height          =   315
         Left            =   1800
         TabIndex        =   31
         Top             =   1560
         Width           =   2055
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   6360
         TabIndex        =   28
         Top             =   1560
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcProducto 
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5400
         TabIndex        =   24
         Top             =   480
         Width           =   765
      End
      Begin VB.CheckBox chkFApertura 
         Alignment       =   1  'Right Justify
         Caption         =   "F. Apert"
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
         TabIndex        =   5
         Top             =   2280
         Width           =   1280
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   8295
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   76414977
            CurrentDate     =   39681
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   5880
            TabIndex        =   15
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   76414977
            CurrentDate     =   39681
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta:"
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
            Left            =   4080
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Desde:"
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtAImporte 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtTasa 
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
      End
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Nro Cuenta:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   6240
         TabIndex        =   23
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label Label6 
         Caption         =   "Direccion:"
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
         Left            =   4080
         TabIndex        =   35
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "A:"
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
         Left            =   1440
         TabIndex        =   30
         Top             =   1980
         Width           =   195
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7050
         TabIndex        =   25
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   4080
         TabIndex        =   22
         Top             =   900
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   360
         TabIndex        =   11
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de Interés:"
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
         Left            =   4080
         TabIndex        =   10
         Top             =   1980
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
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
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1620
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmGenerarCarta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmGenerarCarta
'** Descripción : Formulario para Generar las Cartas de Circularización .
'** Creación : MAVM, 20080814 8:58:15 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Dim lsmensaje As String
Dim NroCta As String
Dim oGen As COMDConstSistema.DCOMGeneral
Dim sOpeCodAho As String

Private Sub chkFApertura_Click()
    If chkFApertura.value = 1 Then
        Frame3.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        DTPicker1.Visible = True
        DTPicker2.Visible = True
    Else
        Frame3.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        DTPicker1.Visible = False
        DTPicker2.Visible = False
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = 1 Then
        TxtAgencia.Text = ""
        lblAgencia.Caption = ""
        chkTodos.value = 1
    Else
        chkTodos.value = 0
    End If
End Sub

Private Sub Command1_Click()
    If chkTodos.value <> 0 Or TxtAgencia.Text <> "" Then
        BuscarDatos
    Else
        MsgBox "Debe Ingresar la Agencia", vbCritical, "Aviso"
    End If
End Sub

Private Sub Command2_Click()
If MsgBox("Esta Seguro de Imprimir los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    If lsmensaje = "" Then
        NroCta = dgBuscar.Columns(3).Text
        ImprimeFormatoCarta (1)
    End If
End If
End Sub

Public Function ImprimeFormatoCarta(tipo As Integer) As String ' Tipo = 1 : Imprime registro seleccionado / Tipo = 2 : Imprime todos los registros //By BRGO 26/10/2010
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim Funcionario As String
    Dim Plantilla As String
    Dim i As Integer
    Dim NumFilas As Integer
    
    On Error GoTo ErrGeneraRepo
    Screen.MousePointer = 11
    lsmensaje = ""
    '*** By BRGO 26/10/2010 ***************************************
    If tipo = 1 Then
        NumFilas = 1
    Else
        Set rs = objCOMNAuditoria.ObtenerDatosCtaAhorro(IIf(txtCuenta.NroCuenta = "109232", "", txtCuenta.NroCuenta), txtCodigo.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcProducto.BoundText = "0", "", dcProducto.BoundText), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), txtDeImporte.Text, txtAImporte.Text, txtTasa.Text, IIf(chkFApertura.value = 0, "", DTPicker1.value), IIf(chkFApertura.value = 0, "", DTPicker2.value), lsmensaje)
        NumFilas = rs.RecordCount
    End If
    
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    'Add By Gitu 23-02-2010
    'Para imprimir las cartas de circulacion con un modelo diferente al de auditoria
    If sOpeCodAho = "" Then
        Plantilla = App.path & "\FormatoCarta\ForCartaCirculacionAhorro.doc"
    Else
        Plantilla = App.path & "\FormatoCarta\ForCartaCirculacionAhorro2.doc"
    End If
     
    wAppSource.Documents.Open Filename:=Plantilla
    wAppSource.ActiveDocument.Content.Copy
    wApp.Documents.Add
    
    For i = 1 To NumFilas
        If tipo = 2 Then
            NroCta = rs.Fields("cCtaCod")
        End If
        Set rs1 = objCOMNAuditoria.ObtenerDatosXNroCta(NroCta)
        Set Rs2 = objCOMNAuditoria.ObtenerDatosFuncionario
        Funcionario = PstaNombre(Rs2.Fields("cPersNombre"))
        
        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        With wApp.Selection.Find
                .Text = "<<NroCta>>"
                .Replacement.Text = rs1.Fields("cCtaCod")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With
        
        With wApp.Selection.Find
                .Text = "<<TPro>>"
                .Replacement.Text = rs1.Fields("cTipoAho")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With
        
        With wApp.Selection.Find
                .Text = "<<Mone>>"
                .Replacement.Text = rs1.Fields("sMoneda")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With
        
        With wApp.Selection.Find
                .Text = "<<Mont>>"
                .Replacement.Text = Format(rs1.Fields("nSaldoDisp"), "#,##0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With
        
        With wApp.Selection.Find
                .Text = "<<NomPers>>"
                .Replacement.Text = rs1.Fields("cPersNombre")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With
        
        With wApp.Selection.Find
                .Text = "<<DirecPers>>"
                .Replacement.Text = rs1.Fields("sDireccion")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With

        Dim liPosicion As Integer
        Dim lsCiudad As String
        lsCiudad = Trim(rs1.Fields("cUbiGeoDescripcion"))
        liPosicion = InStr(lsCiudad, "(")
        
        If liPosicion > 0 Then
            lsCiudad = Left(lsCiudad, liPosicion - 1)
        End If
        With wApp.Selection.Find
            .Text = "<<Fecha>>"
            .Replacement.Text = lsCiudad & ": " & Format(Now, "Long Date")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        If sOpeCodAho = "" Then
            With wApp.Selection.Find
                    .Text = "<<Fun>>"
                    .Replacement.Text = Funcionario
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
            End With
        End If
        If tipo = 2 Then
            rs.MoveNext
        End If
    Next
    If tipo = 2 Then
        rs.Close: Set rs = Nothing
    End If
    rs1.Close: Set rs1 = Nothing
    Rs2.Close: Set Rs2 = Nothing
    Screen.MousePointer = 0
    
    wAppSource.ActiveDocument.Close
    wApp.ActiveDocument.CopyStylesFromTemplate (Plantilla)
    wApp.ActiveDocument.SaveAs (App.path & "\Spooler\FCC_" & NroCta & "_" & Replace(Left(Time, 5), ":", "") & ".doc")
    wApp.Visible = True
    Set wAppSource = Nothing
    Set wApp = Nothing
    Exit Function
    '*** End BRGO **************************************************
ErrGeneraRepo:
        Screen.MousePointer = 0
        wAppSource.ActiveDocument.Close
        MsgBox "Error en frmGenerarCarta.ImprimeFormatoCarta " & Err.Description, vbInformation, "Aviso"
End Function

Private Sub Command3_Click()
    If MsgBox("Esta Seguro de Imprimir los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    
    ' Comentado por BRGO 26/10/2010
    
    '        Dim rs As New ADODB.Recordset
    '        Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    '        Dim i As Integer
    '        lsmensaje = ""
    '        Set rs = objCOMNAuditoria.ObtenerDatosCtaAhorro(IIf(txtCuenta.NroCuenta = "109232", "", txtCuenta.NroCuenta), txtCodigo.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcProducto.BoundText = "0", "", dcProducto.BoundText), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), txtDeImporte.Text, txtAImporte.Text, txtTasa.Text, IIf(chkFApertura.value = 0, "", DTPicker1.value), IIf(chkFApertura.value = 0, "", DTPicker2.value), lsmensaje)
    '
    '        For i = 1 To rs.RecordCount
    '            NroCta = rs.Fields("cCtaCod")
                ImprimeFormatoCarta (2)
    '            rs.MoveNext
    '        Next i
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.value = Date
    DTPicker2.value = Date
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Me.TxtAgencia.rs = oCons.getAgencias(, , True)
    chkTodos.value = 1
    CargarNroCta
    CargarProducto
    CargarMoneda
    cboMoneda.SelText = "Todos"
       
    Frame3.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    
End Sub

Private Sub CargarMoneda()
    cboMoneda.AddItem "Todos", 0
    cboMoneda.AddItem "SOLES", gMonedaNacional
    cboMoneda.AddItem "DOLARES", gMonedaExtranjera
End Sub

Private Sub CargarProducto()
    Dim rsProducto As ADODB.Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsProducto = objCOMNAuditoria.ObtenerProductos
    Set dcProducto.RowSource = rsProducto
    dcProducto.BoundColumn = "nConsValor"
    dcProducto.ListField = "cConsDescripcion"
    Set objCOMNAuditoria = Nothing
    Set rsProducto = Nothing
    dcProducto.BoundText = 0
End Sub

Public Sub CargarNroCta()
    Dim nProducto As String
    nProducto = gCapAhorros
    txtCuenta.NroCuenta = ""
    txtCuenta.Prod = Trim(nProducto)
    txtCuenta.EnabledProd = False
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledCMAC = False
End Sub

Public Sub BuscarDatos()
    Dim rs As Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ObtenerDatosCtaAhorro(IIf(txtCuenta.NroCuenta = "109232", "", txtCuenta.NroCuenta), txtCodigo.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcProducto.BoundText = "0", "", dcProducto.BoundText), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), txtDeImporte.Text, txtAImporte.Text, txtTasa.Text, IIf(chkFApertura.value = 0, "", Format(DTPicker1.value, "yyyymmdd")), IIf(chkFApertura.value = 0, "", Format(DTPicker2.value, "yyyymmdd")), lsmensaje)
    
    If lsmensaje = "" Then
    lblMensaje.Visible = False
    dgBuscar.Visible = True
    Set dgBuscar.DataSource = rs
    dgBuscar.Refresh
    Screen.MousePointer = 0
    dgBuscar.SetFocus
    Command2.Visible = True
    Command3.Visible = True

    Else
    Set dgBuscar.DataSource = Nothing
    dgBuscar.Refresh
    lblMensaje.Visible = True
    lblMensaje.Caption = "No Existen Datos"
    dgBuscar.Visible = False
    Command2.Visible = False
    Command3.Visible = False

    End If
    
    Set rs = Nothing
    Set objCOMNAuditoria = Nothing

End Sub

Private Sub TxtAgencia_EmiteDatos()
    Set oGen = New COMDConstSistema.DCOMGeneral
    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
    chkTodos.value = 0
End Sub

Private Sub txtCodigo_EmiteDatos()
If txtCodigo.Text <> "" Then
    Call CargarCtaAhorroXCliente(txtCodigo.Text)
    Set dgBuscar.DataSource = Nothing
    dgBuscar.Refresh
    Command2.Visible = False
    Command3.Visible = False
End If
End Sub

Public Sub CargarCtaAhorroXCliente(ByVal CodPer As String)
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim Rs2 As ADODB.Recordset
    Dim lsmensaje As String
    
    Set Rs2 = objCOMNAuditoria.ObtenerDatosCtaAhorroXCliente(CodPer, lsmensaje)
    If lsmensaje = "" Then
    txtCliente.Text = Rs2("cPersNombre")
    txtDireccion.Text = Rs2("sDireccion")
    Else
    MsgBox lsmensaje, vbCritical, "Aviso"
    txtCodigo.Text = ""
    txtCliente.Text = ""
    End If
End Sub

Private Sub txtDeImporte_LostFocus()
    FormatoMoneda
End Sub

Private Sub txtAImporte_LostFocus()
    FormatoMoneda
End Sub

Sub FormatoMoneda()
    If Len(txtDeImporte.Text) > 0 Then
    txtDeImporte.Text = Format(txtDeImporte, "#,##0.00")
    End If
    If Len(txtAImporte.Text) > 0 Then
    txtAImporte.Text = Format(txtAImporte, "#,##0.00")
    End If
End Sub
'Add By Gitu
Public Sub Inicia(Optional ByVal psOpeCod As String)
    sOpeCodAho = psOpeCod
    Me.Show 1
End Sub
