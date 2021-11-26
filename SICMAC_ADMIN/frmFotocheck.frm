VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFotocheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Fotocheck"
   ClientHeight    =   6465
   ClientLeft      =   2385
   ClientTop       =   2445
   ClientWidth     =   10545
   Icon            =   "frmFotocheck.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin Sicmact.Usuario Usuario1 
      Left            =   2640
      Top             =   6480
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox txtbuscar 
      Height          =   345
      Left            =   3990
      TabIndex        =   18
      Top             =   5070
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir"
      Height          =   1140
      Left            =   8205
      TabIndex        =   14
      Top             =   4770
      Width           =   2055
      Begin VB.CheckBox chkSelecion 
         Caption         =   "Sólo Pers. Selecc."
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   285
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chksegundo 
         Caption         =   "Segundo Nombre"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   795
         Width           =   1695
      End
      Begin VB.CheckBox chkPrintCargo 
         Caption         =   "No Imprimir Cargo"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   510
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdCodificar 
      Caption         =   "&Codificar Tarjetas"
      Height          =   420
      Left            =   6540
      TabIndex        =   12
      Top             =   5925
      Width           =   1350
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   420
      Left            =   5205
      TabIndex        =   11
      Top             =   5925
      Width           =   1350
   End
   Begin VB.PictureBox Usuario 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00000080&
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8835
      TabIndex        =   2
      Top             =   5925
      Width           =   1395
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar Datos"
      Height          =   405
      Left            =   3930
      TabIndex        =   1
      Top             =   5940
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid dgPers 
      Height          =   4635
      Left            =   3990
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8176
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
   Begin VB.PictureBox picfotocheck 
      Height          =   6195
      Left            =   0
      ScaleHeight     =   108.215
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   65.352
      TabIndex        =   3
      Top             =   120
      Width           =   3765
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   225
         Left            =   1575
         TabIndex        =   10
         Top             =   4290
         Width           =   420
      End
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cargo"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   225
         Left            =   1530
         TabIndex        =   9
         Top             =   4545
         Width           =   510
      End
      Begin VB.Label lblDNI 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI : 12321321"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1035
         TabIndex        =   8
         Top             =   3930
         Width           =   1500
      End
      Begin VB.Label lblApellido 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APELLIDOS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1110
         TabIndex        =   7
         Top             =   3540
         Width           =   1350
      End
      Begin VB.Label lblNombres 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1065
         TabIndex        =   6
         Top             =   3075
         Width           =   1440
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
      Begin VB.Image PicFoto 
         Height          =   2025
         Left            =   1020
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1560
      End
      Begin VB.Image ImgFondo 
         Height          =   6075
         Left            =   -30
         Picture         =   "frmFotocheck.frx":030A
         Stretch         =   -1  'True
         Top             =   15
         Width           =   3720
      End
   End
   Begin MSComDlg.CommonDialog cmdlOpen 
      Left            =   7980
      Top             =   6030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Buscar :"
      Height          =   195
      Left            =   4005
      TabIndex        =   13
      Top             =   4830
      Width           =   585
   End
   Begin VB.Image imgAtrazLoc 
      Height          =   2280
      Left            =   885
      Picture         =   "frmFotocheck.frx":7D80C
      Stretch         =   -1  'True
      Top             =   7065
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Image imgFondoLoc 
      Height          =   3045
      Left            =   -495
      Picture         =   "frmFotocheck.frx":83A64
      Stretch         =   -1  'True
      Top             =   7005
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Image imgAtraz 
      Height          =   3495
      Left            =   0
      Picture         =   "frmFotocheck.frx":8D4CE
      Stretch         =   -1  'True
      Top             =   7065
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
CargaDatos
End Sub

Private Sub cmdCodificar_Click()
cmdlOpen.ShowPrinter
If rs Is Nothing Then Exit Sub
If rs.EOF And rs.BOF Then Exit Sub
If chkSelecion.value = 0 Then
    rs.MoveFirst
    Do While Not rs.EOF
        If PicFoto.Picture <> 0 Then
            CodificarFotoCheck rs!cRHCod, rs!cPersCod, Date
        End If
        rs.MoveNext
        DoEvents
    Loop
Else
    If PicFoto.Picture <> 0 Then
        CodificarFotoCheck rs!cRHCod, rs!cPersCod, Date
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
            If Left(rs!cRHCod, 1) = "L" Then
                ImprimeFotoCheckLoc   ''rs!cRHCod, rs!cPersCod, Date
            Else
                ImprimeFotoCheck  ''rs!cRHCod, rs!cPersCod, Date
            End If
        End If
        rs.MoveNext
        DoEvents
    Loop
Else
    If PicFoto.Picture <> 0 Then
        If Left(rs!cRHCod, 1) = "L" Then
                ImprimeFotoCheckLoc   ''rs!cRHCod, rs!cPersCod, Date
            Else
                ImprimeFotoCheck  ''rs!cRHCod, rs!cPersCod, Date
            End If
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
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

Printer.PaintPicture ImgFondo.Picture, 0, 0, 54, 85.7
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
Sub CargaDatos()
Dim oRh As DActualizaDatosRRHH
Set oRh = New DActualizaDatosRRHH
Set rs = oRh.GetPersonal
Set dgPers.DataSource = rs
dgPers.Refresh
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not pRecordset.EOF And Not pRecordset.BOF Then
    CargaDatosFotochek pRecordset!cPersCod, pRecordset!cRHCargoDescripcion, pRecordset!cPersNombre, pRecordset!cAreaDescripcion
End If
End Sub

Public Sub ImprimeFotoCheckLoc()
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnFontSize As Integer

Printer.Orientation = vbPRORLandscape
Printer.ScaleMode = 6 'Vb Milimetros
Printer.CurrentX = 0 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 0 'Arriba, para la orientación de los driver's
Printer.fontname = "System"

Printer.PaintPicture imgFondoLoc.Picture, 0, 0, 85.7, 54
Printer.PaintPicture PicFoto.Picture, 5, 13, 23, 30


Printer.fontname = "Arial"
Printer.Fontsize = 7
Printer.ForeColor = &H80&
Printer.CurrentX = 46 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 4 'Arriba, para la orientación de los driver's
Printer.Print "CAJA MUNICIPAL"


Printer.fontname = "Arial"
Printer.Fontsize = 7
Printer.ForeColor = &H80&
Printer.CurrentX = 46 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 7 'Arriba, para la orientación de los driver's
Printer.Print "DE AHORRO Y CREDITO DE"

Printer.fontname = "Arial"
Printer.Fontsize = 7
'Printer.FontBold = True
Printer.ForeColor = &H80&
Printer.CurrentX = 70 'Izquierda, para la orientación de los driver's
Printer.CurrentY = 10 'Arriba, para la orientación de los driver's
Printer.Print "ICA S.A."

Printer.fontname = "Arial Narrow"
Printer.Fontsize = 12
Printer.FontBold = True
Printer.ForeColor = &H80&
Printer.CurrentX = 38 'GetCoordCentral(Trim(lblNombres))  'Izquierda, para la orientación de los driver's
Printer.CurrentY = 25  'Arriba, para la orientación de los driver's
Printer.Print Trim(lblNombres)

Printer.fontname = "Arial Narrow"
Printer.Fontsize = 12
Printer.FontBold = True
Printer.ForeColor = &H80&
Printer.CurrentX = 38 'GetCoordCentral(Trim(lblApellido))
Printer.CurrentY = 30 'Arriba, para la orientación de los driver's
Printer.Print Trim(lblApellido)

Printer.FontBold = False

'Printer.FontName = "Arial Narrow"
'Printer.FontSize = 9
'Printer.FontBold = True
'Printer.ForeColor = RGB(77, 77, 77)
'Printer.CurrentX = GetCoordCentral(Trim(lblDNI))
'Printer.CurrentY = 58 'Arriba, para la orientación de los driver's
'Printer.Print Trim(lblDNI)
'Printer.FontItalic = False

If Me.chkPrintCargo.value = 0 Then

    'Printer.FontName = "Arial Narrow"
    'Printer.FontSize = 9
    'Printer.FontBold = True
    'Printer.ForeColor = RGB(74, 72, 106)
    'Printer.CurrentX = 35 ' GetCoordCentral(Trim(lblArea))
    'Printer.CurrentY = 63 'Arriba, para la orientación de los driver's
    'Printer.Print Trim(lblArea)
    
    Printer.fontname = "Arial Narrow"
    Printer.Fontsize = 9
    Printer.FontBold = True
    Printer.ForeColor = &H80&
    Printer.CurrentX = 38 'GetCoordCentral(Trim(lblCargo))
    Printer.CurrentY = 35 'Arriba, para la orientación de los driver's
    Printer.Print Trim(lblCargo)
End If

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
