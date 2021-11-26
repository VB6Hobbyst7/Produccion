VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGenerarCartaCredito 
   Caption         =   "Generar Carta de Circularización - Créditos"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "frmGenerarCartaCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   9000
      TabIndex        =   23
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdBuscar 
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
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtAImporte 
         Height          =   315
         Left            =   6480
         TabIndex        =   30
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1440
         TabIndex        =   29
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtDeImporte 
         Height          =   315
         Left            =   6480
         TabIndex        =   28
         Top             =   1560
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo dcAnalista 
         Height          =   315
         Left            =   5280
         TabIndex        =   27
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   8240
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
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
            Format          =   69861377
            CurrentDate     =   39681
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   5880
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
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
            Format          =   69861377
            CurrentDate     =   39681
         End
         Begin VB.Label Label2 
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkFApertura 
         Alignment       =   1  'Right Justify
         Caption         =   "F. Desem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Width           =   1275
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5280
         TabIndex        =   9
         Top             =   480
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcTipoCredito 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   503
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
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   6120
         TabIndex        =   17
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
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   420
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   741
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   32
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label12 
         Caption         =   "A:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   31
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Analista:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6960
         TabIndex        =   18
         Top             =   480
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   8775
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
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
            DataField       =   "cFDesembolso"
            Caption         =   "F. Desembolso"
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
            DataField       =   "cCtaCod"
            Caption         =   "Nro Credito"
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
            DataField       =   "cTipoProd"
            Caption         =   "Tipo Credito"
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
            DataField       =   "Cliente"
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
            DataField       =   "cAnalista"
            Caption         =   "Analista"
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
            DataField       =   "cEstado"
            Caption         =   "Estado"
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
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1995.024
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
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   9000
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdITodos 
         Caption         =   "I. Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGenerarCartaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmGenerarCartaCredito
'** Descripción : Formulario para Generar las Cartas de Circularización .
'** Creación : MAVM, 20080818 8:58:15 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Dim lsmensaje As String
Dim NroCta As String
Dim oGen As COMDConstSistema.DCOMGeneral

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
        CargarAnalistas TxtAgencia.Text
    Else
        chkTodos.value = 0
    End If
End Sub

Private Sub cmdBuscar_Click()
    If chkTodos.value <> 0 Or TxtAgencia.Text <> "" Then
        BuscarDatos
    Else
        MsgBox "Debe Ingresar la Agencia", vbCritical, "Aviso"
    End If
End Sub

Private Sub cmdImprimir_Click()
    If MsgBox("Esta Seguro de Imprimir los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        If lsmensaje = "" Then
            NroCta = dgBuscar.Columns(1).Text
            ImprimeFormatoCarta
        End If
    End If
End Sub

Private Sub cmdITodos_Click()
    If MsgBox("Esta Seguro de Imprimir los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Dim rs As New ADODB.Recordset
        Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
        Dim i As Integer
        lsmensaje = ""
        'Se Agregaron Parametros de Moneda, y Rangos de Desembolso
        Set rs = objCOMNAuditoria.ObtenerDatosCredito(IIf(Len(txtCuenta.NroCuenta) <= 5, "", txtCuenta.NroCuenta), txtCodigo.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcTipoCredito.BoundText = "305", dcTipoCredito.BoundText, Mid(dcTipoCredito.BoundText, 1, 1)), IIf(dcAnalista.BoundText = "0", "0", (Mid(dcAnalista.Text, Len(dcAnalista.Text) - 3, Len(dcAnalista.Text)))), IIf(chkFApertura.value = 0, "", Format(DTPicker1.value, "yyyymmdd")), IIf(chkFApertura.value = 0, "", Format(DTPicker2.value, "yyyymmdd")), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), txtDeImporte.Text, txtAImporte.Text, lsmensaje)
        For i = 1 To rs.RecordCount
            NroCta = rs.Fields("cCtaCod")
            ImprimeFormatoCarta
            rs.MoveNext
        Next i
    End If
End Sub

Public Function ImprimeFormatoCarta() As String
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Funcionario As String

Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim sArchivo As String

Set rs1 = objCOMNAuditoria.ObtenerCreditoXNroCta(NroCta)
Set rs2 = objCOMNAuditoria.ObtenerDatosFuncionario
Funcionario = PstaNombre(rs2.Fields("cPersNombre"))

Set oWord = CreateObject("Word.Application")
oWord.Visible = False
Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\ForCartaCirculacionCredito.doc")

sArchivo = App.path & "\FormatoCarta\FCCCredito_" & NroCta & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
oDoc.SaveAs (sArchivo)

With oWord.Selection.Find
        .Text = "<<NroCta>>"
        .Replacement.Text = rs1.Fields("cCtaCod")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<Fun>>"
        .Replacement.Text = Funcionario
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<TPro>>"
        .Replacement.Text = rs1.Fields("cTipoProd")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<Mone>>"
        .Replacement.Text = rs1.Fields("cMoneda")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<Mont>>"
        .Replacement.Text = Format(rs1.Fields("nMontoApr"), "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<NomPers>>"
        .Replacement.Text = rs1.Fields("Cliente")
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
        
With oWord.Selection.Find
        .Text = "<<Fecha>>"
        .Replacement.Text = lsCiudad & ": " & Format(Now, "Long Date")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<SCap>>"
        .Replacement.Text = Trim(rs1.Fields("nSaldoCap"))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
        .Text = "<<F>>"
        .Replacement.Text = gdFecData 'gdFecSis
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
End With

oDoc.Close
Set oDoc = Nothing

Set oWord = CreateObject("Word.Application")
oWord.Visible = True

Set oDoc = oWord.Documents.Open(sArchivo)
'oWord.PrintOut
oWord.Visible = True
Set oDoc = Nothing
Set oWord = Nothing
End Function

Private Sub Form_Load()
    DTPicker1.value = Date
    DTPicker2.value = Date
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Me.TxtAgencia.rs = oCons.getAgencias(, , True)
    chkTodos.value = 1
    CargarNroCta
    CargarTipoCredito
    CargarMoneda
    CargarAnalistas TxtAgencia.Text
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

Private Sub CargarAnalistas(ByVal sAgencia As String)
    Dim rsAnalista As New ADODB.Recordset
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsAnalista.DataSource = objCOMNAuditoria.DarAnalista(sAgencia)
    dcAnalista.BoundColumn = "cPersCod"
    dcAnalista.DataField = "cPersCod"
    Set dcAnalista.RowSource = rsAnalista
    dcAnalista.ListField = "cPersNombre"
    dcAnalista.BoundText = 0
End Sub

Private Sub CargarTipoCredito()
    Dim rsProducto As ADODB.Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsProducto = objCOMNAuditoria.ObtenerTipoCredito
    Set dcTipoCredito.RowSource = rsProducto
    dcTipoCredito.BoundColumn = "nConsValor"
    dcTipoCredito.ListField = "cConsDescripcion"
    Set objCOMNAuditoria = Nothing
    Set rsProducto = Nothing
    dcTipoCredito.BoundText = 0
End Sub

Public Sub CargarNroCta()
    txtCuenta.NroCuenta = ""
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Age = gsCodAge
End Sub

Public Sub BuscarDatos()
    Dim rs As Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ObtenerDatosCredito(IIf(Len(txtCuenta.NroCuenta) <= 5, "", txtCuenta.NroCuenta), txtCodigo.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcTipoCredito.BoundText = "305", dcTipoCredito.BoundText, Mid(dcTipoCredito.BoundText, 1, 1)), IIf(dcAnalista.BoundText = "0", "0", (Mid(dcAnalista.Text, Len(dcAnalista.Text) - 3, Len(dcAnalista.Text)))), IIf(chkFApertura.value = 0, "", Format(DTPicker1.value, "yyyymmdd")), IIf(chkFApertura.value = 0, "", Format(DTPicker2.value, "yyyymmdd")), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), txtDeImporte.Text, txtAImporte.Text, lsmensaje)
        If lsmensaje = "" Then
            lblMensaje.Visible = False
            dgBuscar.Visible = True
            Set dgBuscar.DataSource = rs
            dgBuscar.Refresh
            Screen.MousePointer = 0
            dgBuscar.SetFocus
            cmdImprimir.Visible = True
            cmdITodos.Visible = True
    
        Else
            Set dgBuscar.DataSource = Nothing
            dgBuscar.Refresh
            lblMensaje.Visible = True
            lblMensaje.Caption = "No Existen Datos"
            dgBuscar.Visible = False
            cmdImprimir.Visible = False
            cmdITodos.Visible = False
        End If
    Set rs = Nothing
    Set objCOMNAuditoria = Nothing
End Sub

Private Sub TxtAgencia_EmiteDatos()
    Set oGen = New COMDConstSistema.DCOMGeneral
    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
    chkTodos.value = 0
    CargarAnalistas TxtAgencia.Text
End Sub


Private Sub txtCodigo_EmiteDatos()
    If txtCodigo.Text <> "" Then
        Call CargarCreditoXCliente(txtCodigo.Text)
        Set dgBuscar.DataSource = Nothing
        dgBuscar.Refresh
        cmdImprimir.Visible = False
        cmdITodos.Visible = False
    End If
End Sub

Public Sub CargarCreditoXCliente(ByVal CodPer As String)
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs2 As ADODB.Recordset
    Dim lsmensaje As String
    
    Set rs2 = objCOMNAuditoria.ObtenerDatosCreditoXCliente(CodPer, lsmensaje)
        If lsmensaje = "" Then
            txtCliente.Text = rs2("Cliente")
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
