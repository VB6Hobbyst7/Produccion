VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapServicioPagoCargaArchivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convenio de Pago - Carga de Archivo"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   Icon            =   "frmCapServicioPagoCargaArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Carga de Información"
      TabPicture(0)   =   "frmCapServicioPagoCargaArchivo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblReferencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRegistros"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FRCargaArchivo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FEBeneficiarios"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSalir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGuardar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtReferencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTotal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtReferencia 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   1
         Top             =   6390
         Width           =   4935
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6360
         Width           =   1095
      End
      Begin SICMACT.FlexEdit FEBeneficiarios 
         Height          =   4095
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7223
         Cols0           =   5
         EncabezadosNombres=   "#-N° DNI-Apellidos y Nombre-Monto S/.-cPersCod"
         EncabezadosAnchos=   "500-1200-5000-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-R-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         ColWidth0       =   495
         RowHeight0      =   300
      End
      Begin VB.Frame FRCargaArchivo 
         Caption         =   "Carga de Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   8415
         Begin VB.TextBox txtConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1030
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox txtEmpresa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1030
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   6015
         End
         Begin VB.CommandButton cmdExaminar 
            Caption         =   "&Examinar"
            Height          =   375
            Left            =   7200
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog dlgArchivo 
            Left            =   7800
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblConvenio 
            Caption         =   "Convenio:"
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
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblEmpresa 
            Caption         =   "Empresa:"
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
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Total S/.:"
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
         Left            =   5400
         TabIndex        =   14
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label lblRegistros 
         Caption         =   "Total Registros: 0"
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
         TabIndex        =   8
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label lblReferencia 
         Caption         =   "Referencia:"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   6450
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCapServicioPagoCargaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPagoCargaArchivo
'*** Descripción : Formulario para cargar la trama de un convenio.
'*** Creación : ELRO el 20130702 04:24:48 PM, según RFC1306270002
'********************************************************************
Option Explicit

Dim fsruta As String
Dim fsCodigoConvenio, fsNombreConvenio, fsNombreAchivo As String

Private Sub cmdExaminar_Click()

fsruta = Empty
dlgArchivo.InitDir = "C:\"
dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
dlgArchivo.ShowOpen
If dlgArchivo.Filename <> Empty Then
    fsruta = dlgArchivo.Filename
    cmdExaminar.Enabled = False
    cmdGuardar.Enabled = False
    LimpiaFlex FEBeneficiarios
    lblRegistros = "Total Registros: 0"
    txtTotal = ""
    fsCodigoConvenio = ""
    fsNombreConvenio = ""
    fsNombreAchivo = dlgArchivo.FileTitle
    txtReferencia = ""
    cargarArchivo
    cmdExaminar.Enabled = True
Else
    MsgBox "No se eligio un archivo.", vbInformation, "Aviso"
    Exit Sub
End If
End Sub

Private Sub cargarArchivo()

If Trim(fsruta) = "" Then Exit Sub

Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim oUCOMPersona As UCOMPersona
Dim rsPersona  As ADODB.Recordset
Dim rsConvenio  As ADODB.Recordset
Dim lsDOI As String
Dim lnPosicion As Integer


Dim fs As Scripting.FileSystemObject

If InStr(Trim(UCase(fsruta)), ".XLS") <> 0 Or InStr(Trim(UCase(fsruta)), ".XLSX") <> 0 Then

    'Variable de tipo Aplicación de Excel
    Dim oExcel As Excel.Application
    Dim lnFila1, lnFila2, lnFilasFormato As Integer
    Dim lbExisteError As Boolean
    
    '***Para verificar la existencia del archivo en la ruta
    Set fs = New Scripting.FileSystemObject
 
    'Una variable de tipo Libro de Excel
    Dim oLibro As Excel.Workbook
    Dim oHoja As Excel.Worksheet

    'creamos un nuevo objeto excel
    Set oExcel = New Excel.Application
     
    lnFilasFormato = 5000
 
    'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
    
    If fs.FileExists(fsruta) Then
        Set oLibro = oExcel.Workbooks.Open(fsruta)
    Else
        MsgBox "No existe el archivo en esta ruta: " & fsruta, vbCritical, "Advertencia"
        Set oHoja = Nothing
        Set oLibro = Nothing
        Set oExcel = Nothing
        Set fs = Nothing
        Exit Sub
    End If
    
    'Hacemos referencia a la Hoja
    Set oHoja = oLibro.Sheets(1)
    
    'Hacemos el Excel Visible
    'oLibro.Visible = False
    lbExisteError = False
    lsDOI = ""
    With oHoja
       fsCodigoConvenio = .Cells(5, 2)
       fsNombreConvenio = Trim(.Cells(7, 2))
        txtEmpresa = .Cells(3, 2)
        txtConvenio = fsCodigoConvenio & " - " & fsNombreConvenio
 
        
        For lnFila1 = 10 To lnFilasFormato
            lsDOI = .Cells(lnFila1, 1)
            If Len(Trim(lsDOI)) = 8 And Trim(.Cells(lnFila1, 2)) <> "" And Trim(.Cells(lnFila1, 3)) <> "" Then
                '***Verifica si la persona esta registrada****
                Set oUCOMPersona = New UCOMPersona
                Set rsPersona = New ADODB.Recordset
                Set rsPersona = oUCOMPersona.devolverDatosPersona(1, lsDOI)
                '***Fin Verifica si la persona esta registrada
                If Not (rsPersona.BOF And rsPersona.EOF) Then
                    FEBeneficiarios.lbEditarFlex = True
                    FEBeneficiarios.SetFocus
                    FEBeneficiarios.AdicionaFila
                    FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 1) = lsDOI
                    FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 2) = .Cells(lnFila1, 2)
                    FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 3) = Format$(.Cells(lnFila1, 3), "##,##0.00")
                    FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 4) = rsPersona!cPersCod
                    FEBeneficiarios.lbEditarFlex = False
                Else
                    .Range("A" & lnFila1, "C" & lnFila1).Interior.Color = RGB(255, 255, 0)
                    .Cells(lnFila1, 4) = "LA PERSONA NO ESTA REGISTRADA EN EL SISTEMA."
                    lbExisteError = True
                End If
            ElseIf Len(Trim(lsDOI)) > 0 And Len(Trim(lsDOI)) <> 8 Then
                .Range("A" & lnFila1, "C" & lnFila1).Interior.Color = RGB(255, 255, 0)
                .Cells(lnFila1, 4) = "EL DNI NO CONTIENE 8 CARACTERES"
                lbExisteError = True
            ElseIf Len(Trim(lsDOI)) > 0 And (Trim(.Cells(lnFila1, 2)) = "" Or Trim(.Cells(lnFila1, 3)) = "") Then
                .Range("A" & lnFila1, "C" & lnFila1).Interior.Color = RGB(255, 255, 0)
                .Cells(lnFila1, 4) = "EXISTE CAMPO(S) VACIO(S)"
                lbExisteError = True
            End If
        Next lnFila1
        Set rsPersona = Nothing
        Set oUCOMPersona = Nothing
    End With
    
If lbExisteError = False Then
    cmdGuardar.Enabled = True
    txtReferencia.SetFocus
    oLibro.Close
    oExcel.Quit
    Set oHoja = Nothing
    Set oLibro = Nothing
    Set oExcel = Nothing
Else
    LimpiaFlex FEBeneficiarios
    txtEmpresa = ""
    txtConvenio = ""
    fsruta = ""
    oExcel.Visible = True
    Set oHoja = Nothing
    Set oLibro = Nothing
    Set oExcel = Nothing
    Exit Sub
End If


ElseIf InStr(Trim(UCase(fsruta)), ".TXT") <> 0 Then
    Dim f As Integer
    Dim str_Linea As String
    Dim lsDatos() As String
    Dim lnlinea As Integer
    
    '***Para verificar la existencia del archivo en la ruta
    Set fs = New Scripting.FileSystemObject
    
    If Not fs.FileExists(fsruta) Then
        MsgBox "No existe el archivo en esta ruta: " & fsruta, vbCritical, "Advertencia"
        Set fs = Nothing
        Exit Sub
    End If
        
    f = FreeFile
    
    lnlinea = 0
    Open fsruta For Input As #f
        'Inserta Detalle de Recaudo Temporal
        Do
            Line Input #f, str_Linea
            lsDatos = Split(str_Linea, "|")
            lnlinea = lnlinea + 1
            If Len(lsDatos(0)) > 8 And lnlinea = 1 Then
                fsCodigoConvenio = lsDatos(0)
                fsNombreConvenio = Trim(lsDatos(2))
                txtEmpresa = lsDatos(1)
                txtConvenio = fsCodigoConvenio & " - " & fsNombreConvenio
            End If
            
            If UBound(lsDatos) = 2 Then
                If Len(Trim(lsDatos(0))) = 8 And lnlinea > 1 Then
                    lsDOI = lsDatos(0)
                    '***Verifica si la persona esta registrada****
                    Set oUCOMPersona = New UCOMPersona
                    Set rsPersona = New ADODB.Recordset
                    Set rsPersona = oUCOMPersona.devolverDatosPersona(1, lsDOI)
                    '***Fin Verifica si la persona esta registrada
                    If Not (rsPersona.BOF And rsPersona.EOF) Then
                        FEBeneficiarios.lbEditarFlex = True
                        FEBeneficiarios.SetFocus
                        FEBeneficiarios.AdicionaFila
                        FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 1) = lsDOI
                        FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 2) = lsDatos(1)
                        FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 3) = Format$(lsDatos(2), "##,##0.00")
                        FEBeneficiarios.TextMatrix(FEBeneficiarios.Row, 4) = rsPersona!cPersCod
                        FEBeneficiarios.lbEditarFlex = False
                    Else
                        MsgBox "En la linea N° " & Format(lnlinea, "0000") & " la persona no esta registrada en el Sistema.", vbOKOnly + vbCritical, "Aviso"
                        LimpiaFlex FEBeneficiarios
                        Close #f
                        If Trim(FEBeneficiarios.TextMatrix(1, 1)) <> "" Then
                            lblRegistros = "Total Registros: " & (FEBeneficiarios.Rows - 1)
                        Else
                            lblRegistros = ""
                        End If
                        Set fs = Nothing
                        Exit Sub
                    End If
                ElseIf Len(Trim(lsDatos(0))) > 0 And Len(Trim(lsDatos(0))) <> 8 And lnlinea > 1 Then
                    MsgBox "En la linea N° " & Format(lnlinea, "0000") & " el DNI no contiene 8 caracteres.", vbOKOnly + vbCritical, "Aviso"
                    LimpiaFlex FEBeneficiarios
                    Close #f
                    If Trim(FEBeneficiarios.TextMatrix(1, 1)) <> "" Then
                        lblRegistros = "Total Registros: " & (FEBeneficiarios.Rows - 1)
                    Else
                        lblRegistros = ""
                    End If
                    Set fs = Nothing
                    Exit Sub
                ElseIf Len(Trim(lsDatos(0))) > 0 And (Len(Trim(lsDatos(1))) = 0 Or Len(Trim(lsDatos(2))) = 0) Then
                    MsgBox "En la linea N° " & Format(lnlinea, "0000") & " existen campo(s) vaccio(s).", vbOKOnly + vbCritical, "Aviso"
                    LimpiaFlex FEBeneficiarios
                    Close #f
                    If Trim(FEBeneficiarios.TextMatrix(1, 1)) <> "" Then
                        lblRegistros = "Total Registros: " & (FEBeneficiarios.Rows - 1)
                    Else
                        lblRegistros = ""
                    End If
                    Set fs = Nothing
                    Exit Sub
                End If
            Else
                MsgBox "En la linea N° " & Format(lnlinea, "0000") & " no tiene la estructura correcta.", vbOKOnly + vbCritical, "Aviso"
                LimpiaFlex FEBeneficiarios
                Close #f
                If Trim(FEBeneficiarios.TextMatrix(1, 1)) <> "" Then
                    lblRegistros = "Total Registros: " & (FEBeneficiarios.Rows - 1)
                Else
                    lblRegistros = ""
                End If
                Set fs = Nothing
                Exit Sub
            End If
        Loop While Not EOF(f)
    Close #f
    cmdGuardar.Enabled = True
    txtReferencia.SetFocus
End If
If Trim(FEBeneficiarios.TextMatrix(1, 1)) <> "" Then
    lblRegistros = "Total Registros: " & (FEBeneficiarios.Rows - 1)
Else
    lblRegistros = ""
End If
txtTotal = Format$(FEBeneficiarios.SumaRow(3), "##,##0.00")
Set fs = Nothing
End Sub

Private Sub cmdGuardar_Click()

If Trim(txtReferencia) = "" Then
    MsgBox "Debe ingresar la referencia del archivo.", vbInformation, "Aviso"
    Exit Sub
End If

If Trim(txtTotal) = "" Then Exit Sub

If CCur(txtTotal) = 0# Then Exit Sub

Dim oNCOMContFunciones As NCOMContFunciones
Set oNCOMContFunciones = New NCOMContFunciones
Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsConvenio As ADODB.Recordset
Set rsConvenio = New ADODB.Recordset
Dim lsPersCod, lsNroSerPag As String
Dim lnPosicion As Integer
Dim lnConfirmar As Long
Dim lsMovNro  As String



lnPosicion = InStr(fsCodigoConvenio, "SP")
lsPersCod = Left(fsCodigoConvenio, lnPosicion - 1)
lsNroSerPag = Right(fsCodigoConvenio, Len(fsCodigoConvenio) - (Len(lsPersCod) + 2))
        
 Set rsConvenio = oNCOMCaptaGenerales.obtenerConvenioVigente(lsPersCod, lsNroSerPag)
lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        
If Not (rsConvenio.BOF And rsConvenio.EOF) Then
    If rsConvenio!cNomSerPag = fsNombreConvenio And CDec(rsConvenio!nSaldoDisponible) > CDec(txtTotal) Then
        lnConfirmar = oNCOMCaptaGenerales.guardarDebitoConvenioServicioPago(rsConvenio!Id_SerPag, UCase(txtReferencia), _
                                                                            fsNombreAchivo, txtTotal, Format(gdFecSis, "yyyyMMdd"), lsMovNro, _
                                                                            FEBeneficiarios.GetRsNew)
        If lnConfirmar > 0 Then
            MsgBox "Se guardo satisfactoriamente el archivo del convenio.", vbInformation, "Aviso"
            LimpiaFlex FEBeneficiarios
            lblRegistros = "Total Registros: 0"
            txtTotal = ""
            fsCodigoConvenio = ""
            fsNombreConvenio = ""
            fsNombreAchivo = ""
            fsruta = ""
            txtReferencia = ""
            txtEmpresa = ""
            txtConvenio = ""
            Set oNCOMContFunciones = Nothing
            Set oNCOMCaptaGenerales = Nothing
            Set rsConvenio = Nothing
        Else
            MsgBox "No se guardo el archivo del convenio.", vbCritical, "Aviso"
            Set oNCOMContFunciones = Nothing
            Set oNCOMCaptaGenerales = Nothing
            Set rsConvenio = Nothing
            Exit Sub
        End If
    ElseIf rsConvenio!cNomSerPag <> fsNombreConvenio Then
        MsgBox "El nombre no coincide con el código del convenio.", vbCritical, "Aviso"
        Set oNCOMContFunciones = Nothing
        Set oNCOMCaptaGenerales = Nothing
        Set rsConvenio = Nothing
        Exit Sub
    ElseIf CDec(rsConvenio!nSaldoDisponible) <= CDec(txtTotal) Then
        MsgBox "El Saldo Disponible de la cuenta es menor o igual que el Saldo Total de la lista del convenio.", vbCritical, "Aviso"
        Set oNCOMContFunciones = Nothing
        Set oNCOMCaptaGenerales = Nothing
        Set rsConvenio = Nothing
        Exit Sub
    End If
Else
    MsgBox "No existe convenio", vbCritical, "Aviso"
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaGenerales = Nothing
    Set rsConvenio = Nothing
    Exit Sub
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

