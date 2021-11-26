VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFotocheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Fotocheck"
   ClientHeight    =   8865
   ClientLeft      =   2385
   ClientTop       =   2445
   ClientWidth     =   12615
   Icon            =   "frmFotocheck.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00000080&
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4305
      TabIndex        =   2
      Top             =   4080
      Width           =   1395
   End
   Begin VB.CommandButton cmdCodificar 
      Caption         =   "&Codificar Tarjetas"
      Height          =   390
      Left            =   2910
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   1515
      TabIndex        =   6
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mostrar"
      Height          =   1095
      Left            =   8040
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
      Begin VB.OptionButton OptPla 
         Caption         =   "Planilla"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   795
         Width           =   1095
      End
      Begin VB.OptionButton OptPrac 
         Caption         =   "Practicantes"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton OptLoc 
         Caption         =   "Locacion"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin Sicmact.Usuario Usuario1 
      Left            =   11040
      Top             =   3480
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox txtbuscar 
      Height          =   345
      Left            =   840
      TabIndex        =   13
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir"
      Height          =   1140
      Left            =   6000
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
      Begin VB.CheckBox chkSelecion 
         Caption         =   "Sólo Pers. Selecc."
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chksegundo 
         Caption         =   "Segundo Nombre"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   795
         Width           =   1695
      End
      Begin VB.CheckBox chkPrintCargo 
         Caption         =   "No Imprimir Cargo"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar Datos"
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid dgPers 
      Height          =   3195
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5636
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdlOpen 
      Left            =   10440
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picfotocheck 
      Height          =   3195
      Left            =   0
      ScaleHeight     =   55.298
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   103.452
      TabIndex        =   3
      Top             =   120
      Width           =   5925
      Begin VB.Label lblDNI 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI : 10033206"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARGO"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3255
         TabIndex        =   17
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3390
         TabIndex        =   16
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label lblApellido 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APELLIDOS"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3120
         TabIndex        =   15
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label lblNombres 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   1290
      End
      Begin VB.Image PicFoto 
         Height          =   1545
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Image imgFondoLoc 
         Height          =   3045
         Left            =   0
         Picture         =   "frmFotocheck.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DE AHORRO Y CREDITO DE ICA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   765
         TabIndex        =   5
         Top             =   5430
         Width           =   1965
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA MUNICIPAL "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   150
         Left            =   765
         TabIndex        =   4
         Top             =   5265
         Width           =   1395
      End
   End
   Begin VB.Image imgAtrazLoc 
      Height          =   2280
      Left            =   480
      Picture         =   "frmFotocheck.frx":A3E5
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   5325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Buscar :"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   585
   End
   Begin VB.Image imgAtrazLoc1 
      Height          =   2280
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Image imgFondoLoc1 
      Height          =   3045
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Image imgAtraz 
      Height          =   3495
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   3990
   End
End
Attribute VB_Name = "frmFotocheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1


Private Sub cmdCargar_Click()
Dim Loc As Integer
Dim Pra As Integer
Dim Pla As Integer

dgPers.ClearFields
Loc = IIf(OptLoc.value = True, 1, 0)
Pra = IIf(OptPrac.value = True, 1, 0)
Pla = IIf(OptPla.value = True, 1, 0)

CargaDatos Loc, Pla, Pra
End Sub

Private Sub cmdCodificar_Click()
cmdlOpen.ShowPrinter
If rs Is Nothing Then Exit Sub
If rs.EOF And rs.BOF Then Exit Sub
If chkSelecion.value = 0 Then
    rs.MoveFirst
    Do While Not rs.EOF
        If PicFoto.Picture <> 0 Then
            CodificarFotoCheck rs!cRhCod, rs!cPersCod, Date
        End If
        rs.MoveNext
        DoEvents
    Loop
Else
    If PicFoto.Picture <> 0 Then
        CodificarFotoCheck rs!cRhCod, rs!cPersCod, Date
    End If
End If
End Sub

Private Sub cmdImprimir_Click()
cmdlOpen.ShowPrinter
If rs Is Nothing Then Exit Sub
If rs.EOF And rs.BOF Then Exit Sub
If chkSelecion.value = 0 Then
    rs.MoveFirst
    Do While Not rs.EOF
        If PicFoto.Picture <> 0 Then
            'If Left(rs!cRhCod, 1) = "L" Then
                ImprimeFotoCheckLoc rs!cRhCod, rs!cPersCod, Date
            'Else
                'ImprimeFotoCheck  ''rs!cRHCod, rs!cPersCod, Date
            'End If
        End If
        rs.MoveNext
        DoEvents
    Loop
Else
    If PicFoto.Picture <> 0 Then
        'If Left(rs!cRhCod, 1) = "L" Then
                ImprimeFotoCheckLoc rs!cRhCod, rs!cPersCod, Date
        '    Else
         '       ImprimeFotoCheck  ''rs!cRHCod, rs!cPersCod, Date
        '    End If
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub





Private Sub Form_Load()
CentraForm Me
Me.Width = 11670
Me.Height = 5025
OptPla.value = 1
End Sub
Public Sub ImprimeFotoCheck()
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnFontSize As Integer
Printer.Orientation = vbPRORPortrait
Printer.ScaleMode = 6 'Vb Milimetros
Printer.CurrentX = 0 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 0 'Arriba, para la orientación de los driver's
Printer.fontname = "System"

'Printer.PaintPicture ImgFondo.Picture, 0, 0, 54, 85.7
Printer.PaintPicture PicFoto.Picture, 14.5, 3, 25, 33

Printer.fontname = "Arial Narrow"
Printer.Fontsize = 13
Printer.FontBold = True
Printer.ForeColor = RGB(77, 77, 77)
Printer.CurrentX = GetCoordCentral(Trim(lblNombres))  'Izquierda, para la orientación de los driver's
Printer.CurrentY = 45  'Arriba, para la orientación de los driver's
Printer.Print Trim(lblNombres)

Printer.fontname = "Arial Narrow"
Printer.Fontsize = 12
Printer.FontBold = True
Printer.ForeColor = RGB(77, 77, 77)
Printer.CurrentX = GetCoordCentral(Trim(lblApellido))
Printer.CurrentY = 51 'Arriba, para la orientación de los driver's
Printer.Print Trim(lblApellido)

Printer.FontBold = False

Printer.fontname = "Arial Narrow"
Printer.Fontsize = 9
Printer.FontBold = True
Printer.ForeColor = RGB(77, 77, 77)
Printer.CurrentX = GetCoordCentral(Trim(lblDNI))
Printer.CurrentY = 58 'Arriba, para la orientación de los driver's
Printer.Print Trim(lblDNI)
Printer.FontItalic = False

If Me.chkPrintCargo.value = 0 Then
    Printer.fontname = "Arial Narrow"
    Printer.Fontsize = 9
    Printer.FontBold = True
    Printer.ForeColor = RGB(74, 72, 106)
    Printer.CurrentX = GetCoordCentral(Trim(lblArea))
    Printer.CurrentY = 63 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblArea)
    
    Printer.fontname = "Arial Narrow"
    Printer.Fontsize = 9
    Printer.FontBold = True
    Printer.ForeColor = RGB(74, 72, 106)
    Printer.CurrentX = GetCoordCentral(Trim(lblCargo))
    Printer.CurrentY = 67 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblCargo)
End If

Printer.fontname = "Arial"
Printer.Fontsize = 6
Printer.ForeColor = vbRed
Printer.CurrentX = 11 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 75 'Arriba, para la orientación de los driver's
Printer.Print "CAJA MUNICIPAL"

Printer.fontname = "Arial"
Printer.Fontsize = 6
Printer.ForeColor = vbRed
Printer.CurrentX = 11 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 77 'Arriba, para la orientación de los driver's
Printer.Print "DE AHORRO Y CREDITO DE"

Printer.fontname = "Arial"
Printer.Fontsize = 6
'Printer.FontBold = True
Printer.ForeColor = vbRed
Printer.CurrentX = 11 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 79 'Arriba, para la orientación de los driver's
Printer.Print "ICA S.A."

Printer.NewPage

Printer.PaintPicture imgAtraz.Picture, 20, 1.5, 34, 85.7

Printer.EndDoc
End Sub
Function GetCoordCentral(ByVal lsTexto As String) As Single
Dim lnAnchoText As Single
Dim lnAnchoTotal As Single
lnAnchoText = Int(Printer.TextWidth(lsTexto))
lnAnchoText = Int(lnAnchoText / 2)

lnAnchoTotal = 27

GetCoordCentral = Int((lnAnchoTotal - lnAnchoText))
End Function
Sub CodificarFotoCheck(ByVal lsCodEmp As String, ByVal lsCodCli As String, ByVal lsFecImp As String)
Printer.Orientation = vbPRORLandscape
Printer.Print "~C0&B 1 " & Trim(lsCodEmp)  'track 1, Alphanumeric allowed
Printer.Print "~C0&B 2 " & Trim(Format(lsFecImp, "yyyyddmm"))  'track 2, numeric only
Printer.Print "~C0&B 3 " & Trim(lsCodCli)  'track 2, numeric only
Printer.EndDoc
End Sub
Sub CargaDatosFotochek(ByVal lsPersCod As String, ByVal lsCargo As String, ByVal lsNomApell As String, ByVal lsArea As String)
Dim lnPos As Integer
Dim lnPos1 As Integer
Dim lsNombre As String
Dim lsApellido As String

Usuario1.DatosPers lsPersCod
lblDNI = "DNI: " + Usuario1.NroDNIUser
lnPos = InStr(1, lsCargo, " ", vbTextCompare)
lnPos1 = InStr(1, lsCargo, "EJECUTIVO", vbTextCompare)
If lnPos1 > 0 Then
    lnPos = InStrRev(lsCargo, " ")
    lblCargo = Mid(lsCargo, 1, lnPos - 1)
Else
    If lnPos > 0 Then
        lblCargo = Mid(lsCargo, 1, lnPos - 1)
    Else
        lblCargo = lsCargo
    End If
End If
lblArea = Replace(Replace(Replace(lsArea, "DEPARTAMENTO", "DPTO."), "DEPARTEMENTO", "DPTO."), "DE ", "")
lnPos = InStr(1, lsNomApell, ",", vbTextCompare)
lblApellido = Mid(lsNomApell, 1, lnPos - 1)
lnPos = InStr(1, lblApellido, "\", vbTextCompare)
If lnPos > 0 Then
    lblApellido = Mid(lblApellido, 1, lnPos - 1)
    lblApellido = Replace(Replace(lblApellido, "\", " DE "), "/", " ")
Else
    lblApellido = lblApellido
    lblApellido = Replace(Replace(lblApellido, "\", " DE "), "/", " ")
End If

lnPos = InStr(1, lsNomApell, ",", vbTextCompare)
lsNombre = Mid(lsNomApell, lnPos + 1, Len(lsNomApell))
lnPos = InStr(1, lsNombre, " ", vbTextCompare)
If lnPos > 0 Then
    If chksegundo.value = 1 Then
        lblNombres = Mid(lsNombre, lnPos + 1, Len(lsNombre))
    Else
        lblNombres = Mid(lsNombre, 1, lnPos - 1)
    End If
Else
    lblNombres = lsNombre
End If
Me.PicFoto.Picture = LoadPicture()
CargaFoto lsPersCod

End Sub
Sub CargaFoto(ByVal psPersCod As String)
Dim oRh As DActualizaDatosRRHH
Dim rs As ADODB.Recordset
Set oRh = New DActualizaDatosRRHH
Set rs = oRh.GetFoto(psPersCod)
If Not (rs.EOF And rs.BOF) Then
    GetPicture rs.Fields(0), PicFoto
End If
rs.Close
End Sub
Sub CargaDatos(pnIndiLoc As Integer, pnIndiPla As Integer, pnIndiPrac As Integer)
Dim oRh As DActualizaDatosRRHH
Set oRh = New DActualizaDatosRRHH
Set rs = oRh.GetPersonal(pnIndiLoc, pnIndiPla, pnIndiPrac)
Set dgPers.DataSource = rs
dgPers.Refresh
End Sub


Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not pRecordset.EOF And Not pRecordset.BOF Then
    CargaDatosFotochek pRecordset!cPersCod, pRecordset!cRHCargoDescripcion, pRecordset!cPersNombre, IIf(Len(pRecordset!cAreaDescripcion) > 18, pRecordset!cAreaDesResumen, pRecordset!cAreaDescripcion)
End If
End Sub

Public Sub ImprimeFotoCheckLoc(ByVal lsCodEmp As String, ByVal lsCodCli As String, ByVal lsFecImp As String)
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnFontSize As Integer

Printer.Orientation = vbPRORLandscape

Printer.Print "~C0&B 1 " & "#" & Trim(lsCodEmp)   'track 1, Alphanumeric allowed
Printer.Print "~C0&B 2 " & Trim(Format(lsFecImp, "yyyyddmm"))  'track 2, numeric only
Printer.Print "~C0&B 3 " & Trim(lsCodCli)  'track 2, numeric only

Printer.ScaleMode = 6 'Vb Milimetros
Printer.CurrentX = 0 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 0 'Arriba, para la orientación de los driver's
Printer.fontname = "System"

Printer.PaintPicture imgFondoLoc.Picture, 0, 0, 85.7, 54
'Printer.PaintPicture PicFoto.Picture, 5, 13, 23, 30
Printer.PaintPicture PicFoto.Picture, 15, 17, 23, 30

'Printer.fontname = "Arial"
'Printer.Fontsize = 7
'Printer.ForeColor = &H80&
'Printer.CurrentX = 46 'Izquierda, para la orientación de los driver's
'Printer.CurrentY = 4 'Arriba, para la orientación de los driver's
'Printer.Print "CAJA MUNICIPAL DE ICA"

'Printer.fontname = "Arial"
'Printer.Fontsize = 7
'Printer.ForeColor = &H80&
'Printer.CurrentX = 46 'Izquierda, para la orientación de los driver's
'Printer.CurrentY = 7 'Arriba, para la orientación de los driver's
'Printer.Print "DE AHORRO Y CREDITO DE"

'Printer.fontname = "Arial"
'Printer.Fontsize = 7
''Printer.FontBold = True
'Printer.ForeColor = &H80&
'Printer.CurrentX = 70 'Izquierda, para la orientación de los driver's
'Printer.CurrentY = 10 'Arriba, para la orientación de los driver's
'Printer.Print "ICA S.A."

Printer.fontname = "Book Antiqua"
Printer.Fontsize = 11
Printer.FontBold = True
Printer.ForeColor = &H0&
Printer.CurrentX = 40   '38 'GetCoordCentral(Trim(lblNombres))  'Izquierda, para la orientación de los driver's
Printer.CurrentY = 17 '25  'Arriba, para la orientación de los driver's
Printer.Print Trim(lblNombres)

Printer.fontname = "Book Antiqua"
Printer.Fontsize = 9
Printer.FontBold = True
Printer.ForeColor = &H0&
Printer.CurrentX = 40 '38 'GetCoordCentral(Trim(lblApellido))
Printer.CurrentY = 23 '30 'Arriba, para la orientación de los driver's
Printer.Print Trim(lblApellido)
Printer.FontBold = False
    
If Me.chkPrintCargo.value = 0 Then
    Printer.fontname = "Book Antiqua"
    Printer.Fontsize = 9
    Printer.FontBold = True
    Printer.ForeColor = &H0&
    Printer.CurrentX = 40 '38 'GetCoordCentral(Trim(lblCargo))
    Printer.CurrentY = 30 '35 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblCargo)
    Printer.fontname = "Book Antiqua"
    Printer.Fontsize = 8
    Printer.FontBold = True
    Printer.ForeColor = &H0&
    Printer.CurrentX = 40
    Printer.CurrentY = 35 '40 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblArea)
End If
    
    Printer.fontname = "Book Antiqua"
    Printer.Fontsize = 8
    Printer.FontBold = True
    Printer.ForeColor = &H0&
    Printer.CurrentX = 40
    Printer.CurrentY = 40 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblDNI)
    Printer.FontItalic = False
Printer.NewPage
Printer.PaintPicture imgAtrazLoc.Picture, 1, 20, 84, 34
Printer.EndDoc
End Sub

Private Sub txtbuscar_Change()
Dim lsCriterio As String
If Len(Trim(txtbuscar.Text)) > 0 Then
   lsCriterio = "cPersNombre LIKE " & txtbuscar.Text & "*"
   'BuscaDato lsCriterio, rs, 0, False
End If
End Sub
