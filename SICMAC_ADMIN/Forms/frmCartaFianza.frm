VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCartaFianza 
   Caption         =   "Carta Fianza"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCartaFianza.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   13680
      TabIndex        =   6
      Top             =   200
      Width           =   990
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   360
      Left            =   12600
      TabIndex        =   5
      Top             =   200
      Width           =   990
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Procesar"
      Height          =   360
      Left            =   11520
      TabIndex        =   4
      Top             =   200
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proveedor"
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   50
      Width           =   10695
      Begin VB.TextBox txtProvNom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Tag             =   "txtnombre"
         Top             =   200
         Width           =   5280
      End
      Begin Sicmact.TxtBuscar txtPerProv 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   200
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSComCtl2.DTPicker txtFechaIni 
         Height          =   285
         Left            =   7440
         TabIndex        =   2
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   80609281
         CurrentDate     =   41586
      End
      Begin MSComCtl2.DTPicker txtFechaFin 
         Height          =   285
         Left            =   9240
         TabIndex        =   3
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   80609281
         CurrentDate     =   41586
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   255
         Left            =   9000
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   255
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   14535
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   360
         Left            =   1920
         TabIndex        =   11
         Top             =   1080
         Width           =   800
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "editar"
         Height          =   360
         Left            =   1080
         TabIndex        =   10
         Top             =   1080
         Width           =   800
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
         Height          =   4935
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   2
         GridLinesFixed  =   1
         GridLinesUnpopulated=   3
         Appearance      =   0
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
      End
   End
End
Attribute VB_Name = "frmCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmCartaFianza
'** Descripción : Registro de las Cartas Fianza creado segun RFC1902190004
'** Creación : TORE, 20190717 10:45:00 AM
'********************************************************************
Option Explicit
Dim oCartaFianza As DCartaFianza
Dim oProv As DLogProveedor
Dim oRS As ADODB.Recordset

Dim xlsAplicacion As Excel.Application
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String, lsFile As String, lsNomHoja As String


Dim lsCodPers As String
Dim lsDescPers As String


Private Sub Form_Load()
    'Centrar el formulario
    With Screen
        Move (.Width - Width) / 2, (.Height - Height) / 2, Width, Height
    End With
    Call FormatoCartasFianza
    txtFechaIni.value = gdFecSis
    txtFechaFin.value = gdFecSis
    cmdActualizar.Visible = False
    cmdVer.Visible = False
End Sub

Private Sub FormatoCartasFianza()
fg.Cols = 12
fg.TextMatrix(0, 0) = " "
fg.TextMatrix(1, 0) = " "

fg.TextMatrix(0, 1) = "CARTAS FIANZA"
fg.TextMatrix(0, 2) = "CARTAS FIANZA"
fg.TextMatrix(0, 3) = "CARTAS FIANZA"
fg.TextMatrix(0, 4) = "CARTAS FIANZA"
fg.TextMatrix(0, 5) = "CARTAS FIANZA"
fg.TextMatrix(0, 6) = "CARTAS FIANZA"
fg.TextMatrix(0, 7) = "CARTAS FIANZA"
fg.TextMatrix(0, 8) = "CARTAS FIANZA"
fg.TextMatrix(0, 9) = "CARTAS FIANZA"
fg.TextMatrix(0, 10) = "CARTAS FIANZA"
fg.TextMatrix(0, 11) = "CARTAS FIANZA"

fg.TextMatrix(1, 1) = ""
fg.TextMatrix(1, 2) = ""
fg.TextMatrix(1, 3) = "N° Carta Fianza"
fg.TextMatrix(1, 4) = "Banco"
fg.TextMatrix(1, 5) = "Descripción"
fg.TextMatrix(1, 6) = "N° Contrato"
fg.TextMatrix(1, 7) = "F. de Emisión"
fg.TextMatrix(1, 8) = "F. Vencimiento"
fg.TextMatrix(1, 9) = "cCodIFI"
fg.TextMatrix(1, 10) = "nDiasVencidos"
fg.TextMatrix(1, 11) = "Estado"


fg.RowHeight(-1) = 285
fg.ColWidth(0) = 400
fg.ColWidth(1) = 800
fg.ColWidth(2) = 800
fg.ColWidth(3) = 1500
fg.ColWidth(4) = 2000
fg.ColWidth(5) = 6000
fg.ColWidth(6) = 1200
fg.ColWidth(7) = 1200
fg.ColWidth(8) = 1200
fg.ColWidth(9) = 0
fg.ColWidth(10) = 0
fg.ColWidth(11) = 1200


fg.MergeCells = flexMergeRestrictColumns
fg.MergeCol(0) = True
fg.MergeCol(5) = True
fg.MergeCol(7) = True
fg.MergeCol(8) = True
fg.MergeCol(11) = True

fg.MergeRow(0) = True
fg.MergeRow(1) = True

fg.RowHeight(0) = 200
fg.RowHeight(1) = 200

fg.ColAlignmentFixed(-1) = flexAlignCenterCenter
fg.ColAlignment(1) = flexAlignCenterCenter
fg.ColAlignment(2) = flexAlignCenterCenter
fg.ColAlignment(3) = flexAlignCenterCenter
fg.ColAlignment(5) = flexAlignLeftCenter
fg.ColAlignment(6) = flexAlignCenterCenter
fg.ColAlignment(7) = flexAlignCenterCenter
fg.ColAlignment(8) = flexAlignCenterCenter
fg.ColAlignment(11) = flexAlignCenterCenter
End Sub

Private Sub fg_EnterCell()
    With fg
        If .col = 1 And Trim(fg.TextMatrix(.RowSel, 3)) <> "" Then
            cmdActualizar.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            cmdActualizar.Visible = True
        End If
        If .col = 2 And Trim(fg.TextMatrix(.RowSel, 3)) <> "" Then
            cmdVer.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            cmdVer.Visible = True
        End If
    End With
End Sub

Private Sub fg_LeaveCell()
    cmdActualizar.Visible = False
    cmdVer.Visible = False
End Sub

Private Sub txtPerProv_EmiteDatos()
    Set oRS = New ADODB.Recordset
    Set oProv = New DLogProveedor
    Set oCartaFianza = New DCartaFianza
    
    Me.txtProvNom.Text = txtPerProv.psDescripcion
    lsCodPers = txtPerProv.psCodigoPersona
    If txtPerProv.psDescripcion <> "" Then
        Set oRS = oProv.GetProveedorAgeRetBuenCont(lsCodPers)
        If oRS.EOF And oRS.BOF Then
            MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub cmdAgregar_Click()
    If txtPerProv.Text <> "" Then
        frmGestionCartaFianza.Inicia "", Trim(txtPerProv.Text), Trim(txtProvNom.Text)
        cmdMostrar_Click
    Else
        MsgBox "Es necesario ingresar la información del proveedor para realizar la acción.", vbInformation, "Aviso"
    End If
End Sub


Private Sub cmdMostrar_Click()
    If txtPerProv.Text <> "" Then
        Set oRS = New ADODB.Recordset
        Set oCartaFianza = New DCartaFianza
        Dim lnItem As Integer
        Dim lbValido As Boolean
        Set oRS = oCartaFianza.ObtenerInfoCartaFianzaCodProv(Trim(txtPerProv.Text), _
                                                                Format(txtFechaIni.value, "yyyyMMdd"), _
                                                                Format(txtFechaFin.value, "yyyyMMdd"))
        fg.Rows = 3
        lnItem = 1
        If Not (oRS.BOF And oRS.EOF) Then
            Do While Not oRS.EOF
                 If lnItem <> 1 Then
                    AdicionaRow fg
                End If
                lnItem = fg.row
                fg.TextMatrix(lnItem, 1) = "Editar"
                fg.TextMatrix(lnItem, 2) = "Ver"
                fg.TextMatrix(lnItem, 3) = oRS!cNroCartaFianza
                fg.TextMatrix(lnItem, 4) = oRS!cIFiNombre
                fg.TextMatrix(lnItem, 5) = oRS!cDescripcion
                fg.TextMatrix(lnItem, 6) = oRS!cNroContrato
                fg.TextMatrix(lnItem, 7) = oRS!dFEmision
                fg.TextMatrix(lnItem, 8) = oRS!dFVencimiento
                fg.TextMatrix(lnItem, 9) = oRS!cCodIfi
                fg.TextMatrix(lnItem, 10) = oRS!nDiasVencidos
                fg.TextMatrix(lnItem, 11) = oRS!cEstadoCF
                
                oRS.MoveNext
            Loop
        Else
             MsgBox "No se encontró Cartas Fianzas del Proveedor", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Es necesario ingresar la información del proveedor para realizar la acción", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdExportar_Click()
    
    If txtPerProv.Text = "" Then
        MsgBox "Es necesario ingresar la información del proveedor para realizar la acción.", vbInformation, "Aviso"
        'txtPerProv.SetFocus
        Exit Sub
    End If

    Set oRS = New ADODB.Recordset
    Set oCartaFianza = New DCartaFianza
    Set xlsAplicacion = New Excel.Application
    Set fs = New Scripting.FileSystemObject
    
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim lsArchivo As String
    Dim lbHojaExiste As Boolean
    Dim lsFormato As String, lsNombreHoja As String
    Dim I As Integer: Dim J As Integer
    
    Dim HoraSis As Variant
    Dim HoraCrea As String
    
    HoraSis = Time
    HoraCrea = CStr(Hour(HoraSis)) & Minute(HoraSis) & Second(HoraSis)
    
    lsNombreHoja = "CartaFianza"
    lsFormato = "FormatoReportCartaFianza"
    
    lsArchivo = "\spooler\" & "CartarFianza_" & gsUser & Format(gdFecSis, "yyyyMMdd") & HoraCrea & ".xls"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFormato & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFormato & ".xlsx")
    Else
        MsgBox "No extiste plantilla para el reporte de formato carta", vbInformation, "Aviso"
        Exit Sub
    End If

     For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNombreHoja Then
            xlHoja1.Activate
            lbHojaExiste = True
            Exit For
       End If
    Next

    If lbHojaExiste = False Then
        MsgBox "La plantilla no es la correcta, no se encontro la hoja de trabajo.", vbInformation, "Aviso"
        Exit Sub
    End If

    Set oRS = oCartaFianza.ObtenerInfoCartaFianzaCodProv(Trim(txtPerProv.Text), _
                                                                Format(txtFechaIni.value, "yyyyMMdd"), _
                                                                Format(txtFechaFin.value, "yyyyMMdd"))

    If (oRS.EOF And oRS.BOF) Then
        MsgBox "No se encontró información de cartas fianzas registrados para el proveedor.", vbInformation, "Aviso"
    
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        
        Exit Sub
    End If
    
    
    xlHoja1.Cells(1, 1) = "Cartas Fianza - " & Space(1) & oRS!cProvNombre

    I = 2
    If Not (oRS.EOF And oRS.BOF) Then
        Do While Not oRS.EOF
            I = I + 1
            xlHoja1.Cells(I, 1) = oRS!cNroCartaFianza
            xlHoja1.Cells(I, 2) = oRS!cNroContrato
            xlHoja1.Cells(I, 3) = oRS!cIFiNombre
            xlHoja1.Cells(I, 4) = oRS!cMoneda
            xlHoja1.Cells(I, 5) = oRS!cDescripcion
            xlHoja1.Cells(I, 6) = oRS!dFEmision
            xlHoja1.Cells(I, 7) = oRS!dFVencimiento
            xlHoja1.Cells(I, 8) = oRS!cEstadoCF
            xlHoja1.Cells(I, 9) = oRS!nDiasVencidos
            
            xlHoja1.Range("A" & Trim(Str(I)) & ":" & "I" & Trim(Str(I))).Borders.LineStyle = 1
            xlHoja1.Range("A" & Trim(Str(I)) & ":" & "I" & Trim(Str(I))).WrapText = True
            
            If oRS.EOF Then
                Exit Do
            End If
            oRS.MoveNext
        Loop
    Else
        MsgBox "No se encontro informacioón de las cartas fianzas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    xlsAplicacion.DisplayAlerts = False
    xlHoja1.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub

Private Sub cmdActualizar_Click()
    With fg
        frmGestionCartaFianza.Inicia Trim(.TextMatrix(.RowSel, 3)), lsCodPers, lsDescPers
    End With
    cmdMostrar_Click
End Sub
Private Sub cmdVer_Click()
     With fg
        frmGestionCartaFianza.VerCF Trim(.TextMatrix(.RowSel, 3))
    End With
End Sub




