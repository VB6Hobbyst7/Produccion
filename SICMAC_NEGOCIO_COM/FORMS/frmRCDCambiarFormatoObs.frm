VERSION 5.00
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmRCDCambiarFormatoObs 
   Caption         =   "Actualizar Archivos"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmRCDCambiarFormatoObs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "Crear Archivo"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtRutaArchivo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdSubirArchivo 
      Caption         =   "Buscar Archivo"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Filtro          =   ""
      Altura          =   0
   End
End
Attribute VB_Name = "frmRCDCambiarFormatoObs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsNomFile As String
Dim fsPathFile As String
Dim fsruta As String
Dim datos As Archivo
Dim i As Integer
Dim tipoArch As Integer
Private Type Archivo
    ntipfor As String * 1
    ntipinf As String * 1
    nnrosec As String * 8
    ccodsbs As String * 10
    ccodempsbs As String * 13
    ccodciu As String * 11
    ccodreg As String * 8
    ctipdtr As String * 1 '0
    cnumdtr As String * 1 '0
    ctipdid As String * 1
    cnumdid As String * 8
    ctipper As String * 1
    ccodres As String * 1
    ccladeu As String * 1
    cMagEmp As String * 1
    caccemp As String * 1
    crelemp As String * 1
    cpainac As String * 2
    csexo As String * 1
    cestciv As String * 1
    cespbla As String * 10
    caprzso As String * 120
    capemat As String * 40
    capecas As String * 40
    cprinom As String * 40
    csegnom As String * 40
    cindrie As String * 1
    csigcom As String * 8
    cnomcli As String * 8
    cdelreg As String * 1 '0
    cli_repe As String * 1 '0
End Type

Private Sub cmdArchivo_Click()
    If Right(fsNomFile, 3) = "OBS" Or Right(fsNomFile, 3) = "obs" Then
    If MsgBox("Se va a generar el archivo excel, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
      Leer_Lineas (fsruta)
      tipoArch = 1
      cmdArchivo.Enabled = True
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSubirArchivo_Click()

    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Filtro = "Archivos Convenio (*.obs)|*.obs"
    
    Me.CdlgFile.altura = 300
    CdlgFile.Show
    
    fsPathFile = CdlgFile.Ruta
    txtRutaArchivo.Text = fsPathFile
    fsruta = fsPathFile
        If fsPathFile <> Empty Then
            For i = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, i, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, i)
                        fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
                        Exit For
                    End If
             Next i
          Screen.MousePointer = 11
          Call cmdArchivo_Click
'          If Right(fsNomFile, 3) = "OBS" Or Right(fsNomFile, 3) = "obs" Then
'            Leer_Lineas (fsruta)
'            tipoArch = 1
'            cmdArchivo.Enabled = True
'          End If
          
        Else
           MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
           Screen.MousePointer = 0
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CentraForm Me
'    cmdArchivo.Enabled = False
End Sub

Public Sub Leer_Lineas(strTextFile As String)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim nCerrar As Integer
    Dim oBarra As clsProgressBar
    Set oBarra = New clsProgressBar
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos
    Dim f As Integer
    Dim fd As String
    Dim p1 As String
    Dim p3 As String
    Dim p2 As Integer
    Dim lineas As Long
    Dim str_Linea As String
    
    'Abre el archivo excel
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "RCDObs"
    lsNomHoja = "rcdobs"
    lsArchivo1 = "\spooler\RCDObs" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    '


    lineas = 0
    f = FreeFile
    nCerrar = 0
    Open strTextFile For Input As #f
    Do 'grabar por cada item el cuerpo
        Line Input #f, str_Linea
        lineas = lineas + 1

        If lineas = 1 And Len(str_Linea) <> 52 Then
                MsgBox "El archivo adjunto NO tiene la cabecera correcta", vbInformation, "Aviso"
                Close #f
                nCerrar = 0
                Exit Sub
        End If
'        ElseIf lineas = 1 And Len(str_Linea) = 52 Then
            xlHoja1.Cells(lineas + 1, 1) = "'" & Mid(str_Linea, 1, 1) '1
            xlHoja1.Cells(lineas + 1, 2) = "'" & Mid(str_Linea, 2, 1) '2
            xlHoja1.Cells(lineas + 1, 3) = "'" & Mid(str_Linea, 3, 8) '11
            xlHoja1.Cells(lineas + 1, 4) = "'" & Mid(str_Linea, 11, 10) '21
            xlHoja1.Cells(lineas + 1, 5) = "'" & Trim(Mid(str_Linea, 21, 13)) '34
            xlHoja1.Cells(lineas + 1, 6) = "'" & Right(Mid(str_Linea, 34, 11), 4) '45
            xlHoja1.Cells(lineas + 1, 7) = "'" & Mid(str_Linea, 45, 8) '53
            xlHoja1.Cells(lineas + 1, 8) = ""
            xlHoja1.Cells(lineas + 1, 9) = "'" & "0"
            xlHoja1.Cells(lineas + 1, 10) = "'" & Right(Mid(str_Linea, 53, 20), 1) '73
            xlHoja1.Cells(lineas + 1, 11) = "'" & Mid(str_Linea, 73, 8) '
            xlHoja1.Cells(lineas + 1, 12) = "'" & Right(Mid(str_Linea, 81, 5), 1)
            xlHoja1.Cells(lineas + 1, 13) = "'" & Mid(str_Linea, 86, 1)
            xlHoja1.Cells(lineas + 1, 14) = "'" & Mid(str_Linea, 87, 1)
            xlHoja1.Cells(lineas + 1, 15) = "'" & Mid(str_Linea, 88, 1)
            xlHoja1.Cells(lineas + 1, 16) = "'" & Mid(str_Linea, 89, 1)
            xlHoja1.Cells(lineas + 1, 17) = "'" & Mid(str_Linea, 90, 1)
            xlHoja1.Cells(lineas + 1, 18) = "'" & Mid(str_Linea, 91, 2)
            xlHoja1.Cells(lineas + 1, 19) = "'" & Trim(Mid(str_Linea, 93, 3))
            xlHoja1.Cells(lineas + 1, 20) = "'" & Trim(Mid(str_Linea, 96, 1))
            xlHoja1.Cells(lineas + 1, 21) = ""
            xlHoja1.Cells(lineas + 1, 22) = "'" & Mid(str_Linea, 117, 120)
            xlHoja1.Cells(lineas + 1, 23) = "'" & Mid(str_Linea, 237, 40)
            xlHoja1.Cells(lineas + 1, 24) = "'" & Mid(str_Linea, 277, 40)
            xlHoja1.Cells(lineas + 1, 25) = "'" & Mid(str_Linea, 317, 40)
            xlHoja1.Cells(lineas + 1, 26) = "'" & Mid(str_Linea, 357, 40)
            xlHoja1.Cells(lineas + 1, 27) = "'" & Mid(str_Linea, 397, 1)
            xlHoja1.Cells(lineas + 1, 28) = "'" & Trim(Mid(str_Linea, 398, 8))
            xlHoja1.Cells(lineas + 1, 29) = "'" & Mid(str_Linea, 425, 8)
            xlHoja1.Cells(lineas + 1, 30) = ""
            xlHoja1.Cells(lineas + 1, 31) = ""
            xlHoja1.Range(xlHoja1.Cells(lineas + 1, 1), xlHoja1.Cells(lineas + 1, 31)).Borders.LineStyle = 1
            'Insertar Datos a Excel
            datos.ntipfor = Mid(str_Linea, 1, 1) '1
            datos.ntipinf = Mid(str_Linea, 2, 1) '2
            datos.nnrosec = Mid(str_Linea, 3, 8) '11
            datos.ccodsbs = Mid(str_Linea, 11, 10) '21
            datos.ccodempsbs = Trim(Mid(str_Linea, 21, 13)) '34
            datos.ccodciu = Right(Mid(str_Linea, 34, 11), 4) '45
            datos.ccodreg = Mid(str_Linea, 45, 8) '53
            datos.ctipdtr = ""
            datos.cnumdtr = "0"
            datos.ctipdid = Right(Mid(str_Linea, 53, 20), 1) '73
            datos.cnumdid = Mid(str_Linea, 73, 8) '
            datos.ctipper = Right(Mid(str_Linea, 81, 5), 1)
            datos.ccodres = Mid(str_Linea, 86, 1)
            datos.ccladeu = Mid(str_Linea, 87, 1)
            datos.cMagEmp = Mid(str_Linea, 88, 1)
            datos.caccemp = Mid(str_Linea, 89, 1)
            datos.crelemp = Mid(str_Linea, 90, 1)
            datos.cpainac = Mid(str_Linea, 91, 2)
            datos.csexo = Trim(Mid(str_Linea, 93, 3))
            datos.cestciv = Trim(Mid(str_Linea, 96, 1))
            datos.cespbla = ""
            datos.caprzso = Mid(str_Linea, 117, 120)
            datos.capemat = Mid(str_Linea, 237, 40)
            datos.capecas = Mid(str_Linea, 277, 40)
            datos.cprinom = Mid(str_Linea, 317, 40)
            datos.csegnom = Mid(str_Linea, 357, 40)
            datos.cindrie = Mid(str_Linea, 397, 1)
            datos.csigcom = Trim(Mid(str_Linea, 398, 8))
            datos.cnomcli = Mid(str_Linea, 425, 8)
            datos.cdelreg = ""
            datos.cli_repe = ""
            
            nCerrar = 0
             'valida el numero de registros a procesar
'        End If
    Loop While Not EOF(f)
    Close #f
    If nCerrar = 1 Then
        rsX.Close
    End If
    oBarra.CloseForm Me
    Set oBarra = Nothing
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub
