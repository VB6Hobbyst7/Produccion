VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmTransferenciaPVS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo PVS"
   ClientHeight    =   1920
   ClientLeft      =   3525
   ClientTop       =   1575
   ClientWidth     =   4860
   Icon            =   "frmTransferenciaPVS.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Generación del Archivo"
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
      Height          =   1125
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   4785
      Begin Spinner.uSpinner uSpinner1 
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Max             =   99
         Min             =   1
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Archivo TXT"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Archivo XLS"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmTransferenciaPVS.frx":030A
         Left            =   1050
         List            =   "frmTransferenciaPVS.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1965
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   3030
         TabIndex        =   0
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Envío:"
         Height          =   195
         Left            =   3120
         TabIndex        =   10
         Top             =   760
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
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
         Left            =   270
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   3420
      TabIndex        =   3
      Top             =   1440
      Width           =   1410
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "Generar"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   1410
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   225
      Left            =   40
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmTransferenciaPVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'** Objeto: frmTransferenciaPVS
'** Fecha : 11/06/2013 07:21 PM
'** Programador : Pedro Acuña - PEAC
'** Descripcion : Genera un archivo para la presentacion del COA
'**               para enviar al programa validador PVS en formato texto.
'****************************************************

Option Explicit
Dim i As Long
Dim rs    As ADODB.Recordset
Dim lsAnio As String
Dim lsMes As String
Dim lsPeriodo As String
Dim sConCoa As String

Private Sub TransfiereComprobanteTXT()

Dim psArchivoAGrabar As String
Dim ArcSal As Integer
Dim sCad As String
Dim vCont As Long
Dim lcVeces As String
Dim obDAgencia As DAgencia
Set obDAgencia = New DAgencia

Dim oCon As DConecta
Set oCon = New DConecta

Dim sSql As String
lsAnio = txtAnio.Text
lsMes = Format(Me.cboMes.ListIndex + 1, "00")
lsPeriodo = lsAnio & lsMes
gsOpeCod = "700101"
lcVeces = Right("00" + Trim(Str(Me.uSpinner1.Valor)), 2)

psArchivoAGrabar = App.path & "\SPOOLER\20103845328001" & lsPeriodo & "." & lcVeces & ".txt"

ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal

'Print #ArcSal, "00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
sCad = ""

On Error GoTo ERROR

Set rs = New ADODB.Recordset
Set rs = obDAgencia.CargaCOAparaPVS(lsPeriodo, gcCtaIGV, gsOpeCod)
Set obDAgencia = Nothing

If rs.EOF Then
    MsgBox "No existen datos para generar los reportes.", vbInformation, "Atención"
    Exit Sub
Else

prgList.Min = 0
prgList.Max = rs.RecordCount
prgList.Visible = True

Do While Not rs.EOF
    Print #1, rs!cCampo
    
    vCont = vCont + 1
    prgList.value = vCont
    
    rs.MoveNext
Loop

Close ArcSal
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & psArchivoAGrabar, vbInformation + vbOKOnly, "Atención"

prgList.Visible = False
End If

Exit Sub
ERROR:
    MsgBox TextErr(Err.Description)
End Sub

Private Sub TransfiereComprobanteXLS()
    Dim liLineas As Integer
    Dim i As Integer
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim lbConexion As Boolean
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim glsArchivo As String
        
    glsArchivo = "COA" & Format(Now, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape
    
    Call ImprimeCOA(xlLibro, xlHoja1, xlAplicacion, "9999", gdFecSis, "99", "0", "CMAC MAYNAS", "Ag. Principal")
            
    xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
    ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo, vbInformation, "Aviso"
    CargaArchivoCOA App.path & "\SPOOLER\" & glsArchivo, App.path & "\SPOOLER\"
End Sub
Public Sub CargaArchivoCOA(lsArchivo As String, lsRutaArchivo As String)
    Dim X As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    X = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If X <= 32 Then
        If X = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf X = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
End Sub

Public Sub ImprimeCOA(xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, xlAplicacion As Excel.Application, psOpeCod As String, pdFecha As Date, psBalanceCate As String, pbSoles As String, psEmpresa As String, psAgenciaCod As String)

Dim lnDivide As Integer
Dim prs As ADODB.Recordset

Dim cnomhoja As String
Dim liLineas As Long
Dim nReg As Integer
Dim lnTipoCambio As Currency
Dim glsArchivo As String
Dim vCont As Long
Dim obDAgencia As DAgencia
Set obDAgencia = New DAgencia
 
'*************INC1709130008****************
lsAnio = txtAnio.Text
lsMes = Format(Me.cboMes.ListIndex + 1, "00")
'*****************Agregado by NAGL 20170928

lsPeriodo = lsAnio & lsMes
gsOpeCod = "700101"
        
    cnomhoja = "COA"
    
    Call ExcelAddHoja(cnomhoja, xlLibro, xlHoja1)

    xlAplicacion.Range("A1:K3").Font.Bold = True
    xlAplicacion.Range("A1:K3").Font.Color = RGB(0, 0, 225)

 Dim dBalance As New NBalanceCont
     
    Set prs = obDAgencia.CargaCOAparaPVS(lsPeriodo, gcCtaIGV, gsOpeCod)
    Set obDAgencia = Nothing
     
    If Not (prs.EOF And prs.BOF) Then
    
        xlHoja1.Cells(1, 1) = "Información para el COA, Periodo: " & lsPeriodo
    
        xlHoja1.Cells(3, 1) = "Num. doc."
        xlHoja1.Cells(3, 2) = "Periodo"
        xlHoja1.Cells(3, 3) = "Fecha Doc."
        xlHoja1.Cells(3, 4) = "Tipo Doc."
        xlHoja1.Cells(3, 5) = "Serie"
        xlHoja1.Cells(3, 6) = "Número"
        xlHoja1.Cells(3, 7) = "Base Imp."
        xlHoja1.Cells(3, 8) = "IGV"
        xlHoja1.Cells(3, 9) = "Tipo Ope."
        xlHoja1.Cells(3, 10) = "Moneda"
        xlHoja1.Cells(3, 11) = "Numero Ref."
        
        liLineas = 3
        
'        prgList.Min = 0
'        prgList.Max = prs.RecordCount
'        prgList.Visible = True
'        vCont = 0
        
        Do While Not prs.EOF
            liLineas = liLineas + 1
            
            xlHoja1.Cells(liLineas, 1) = "'" + Trim(prs!num_doc)
            xlHoja1.Cells(liLineas, 2) = "'" + Trim(prs!periodo)
            xlHoja1.Cells(liLineas, 3) = "'" + Trim(prs!Fecha)
            xlHoja1.Cells(liLineas, 4) = "'" + Trim(prs!tipodoc)
            xlHoja1.Cells(liLineas, 5) = "'" + Trim(prs!Serie)
            xlHoja1.Cells(liLineas, 6) = "'" + Trim(prs!numero)
            xlHoja1.Cells(liLineas, 7) = Format(prs!base_imp, gsFormatoNumeroView)
            xlHoja1.Cells(liLineas, 8) = Format(prs!IGV, gsFormatoNumeroView)
            xlHoja1.Cells(liLineas, 9) = "'" + Trim(prs!tipoope)
            xlHoja1.Cells(liLineas, 10) = "'" + Trim(prs!Moneda)
            xlHoja1.Cells(liLineas, 11) = "'" + Trim(prs!num_ref)
    
'            vCont = vCont + 1
'            prgList.value = vCont
    
            prs.MoveNext
        Loop
    Else
        MsgBox "No existe datos para mostrar.", vbInformation + vbOKOnly, "Atención"
        Exit Sub
   End If
   prs.Close

'   prgList.Visible = False

End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtAnio.SetFocus
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdTransferir_Click()

If Me.Option1.value = False And Me.Option2.value = False Then
    MsgBox "Seleccione el tipo de reporte XLS ó TXT.", vbInformation + vbOKOnly, "Atención"
    Exit Sub
End If
If Me.Option2.value = True Then
    TransfiereComprobanteTXT
ElseIf Me.Option1.value = True Then
    TransfiereComprobanteXLS
End If
   
End Sub

Private Sub Form_Load()
On Error GoTo ConexionErr
    txtAnio.Text = Str(Year(gdFecSis))
    cboMes.ListIndex = Month(gdFecSis) - 1
   Exit Sub
ConexionErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdTransferir.SetFocus
End If
End Sub
