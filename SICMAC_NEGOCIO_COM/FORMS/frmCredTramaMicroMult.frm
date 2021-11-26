VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredTramaMicroMult 
   Caption         =   "Generar Trama (MicroSeguros/Multiriesgos)"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleMode       =   0  'User
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Deposito"
      Height          =   3015
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   6975
      Begin VB.TextBox txtObservaciones 
         Height          =   885
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtNroDocumento 
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipoPago 
         Height          =   315
         ItemData        =   "frmCredTramaMicroMult.frx":0000
         Left            =   1680
         List            =   "frmCredTramaMicroMult.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   3975
      End
      Begin MSMask.MaskEdBox txtMesDepostito 
         Height          =   300
         Left            =   1680
         TabIndex        =   29
         Top             =   960
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Documento:"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   960
         TabIndex        =   31
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Déposito:"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   1000
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Cte. :"
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   650
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pago:"
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame framePrima 
      Caption         =   "Generar Trama"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Frame frmMicroseguro 
         Caption         =   "Tipo Microseguro"
         Height          =   855
         Left            =   4080
         TabIndex        =   39
         Top             =   720
         Width           =   2535
         Begin VB.ComboBox cmbMicroseguros 
            Height          =   315
            ItemData        =   "frmCredTramaMicroMult.frx":01B2
            Left            =   120
            List            =   "frmCredTramaMicroMult.frx":01BC
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.TextBox txtCodPlan 
         Height          =   285
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCodPatrocinador 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   18
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtNumRamo 
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   17
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNumPoliza 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         ItemData        =   "frmCredTramaMicroMult.frx":01DD
         Left            =   1680
         List            =   "frmCredTramaMicroMult.frx":01DF
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3120
         Width           =   2535
      End
      Begin VB.ComboBox cmdTipo 
         Height          =   315
         ItemData        =   "frmCredTramaMicroMult.frx":01E1
         Left            =   1680
         List            =   "frmCredTramaMicroMult.frx":01EB
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cmbTipoDeclaracion 
         Height          =   315
         ItemData        =   "frmCredTramaMicroMult.frx":026D
         Left            =   1680
         List            =   "frmCredTramaMicroMult.frx":0277
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Generación del Archivo"
         Height          =   1575
         Left            =   4320
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
         Begin VB.CommandButton cmdGenerar 
            Caption         =   "&Generar"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton opteExportarTXT 
            Caption         =   "Exportar a TXT"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optExportarExcel 
            Caption         =   "Exportar a Excel"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   1815
         End
      End
      Begin MSMask.MaskEdBox txtMesActual 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   1680
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Format          =   "MM/YYYY"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Plan:"
         Height          =   195
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Patrocinador:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Ramo:"
         Height          =   195
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Producto:"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Poliza:"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lblMesActual 
         Caption         =   "Enero 1900"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Declaracion:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Moneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes a Generar:"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCredTramaMicroMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredTramaMicroMult
'***     Descripcion:      Generar las tramas de Microseguros y Multiriesgo
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     16/05/2012 01:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Dim sMonedaTXT As String
Dim sTipoD As String
Dim sTipoTrama As String
Dim sTipoPago As String
Dim sMes As String
Dim sAno As String
Dim bValidaFec As Boolean
Dim dfechaini As String
Dim dfechafin As String
Dim psArchivoAGrabarC As String
Dim nMonto As Double

Private Sub MicrosegurosEXCEL()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nCerrar As Integer
    Dim i As Integer
    Dim nNumLineas As Integer
    Dim nNumRep As Integer
    Dim nNumMov As Integer
    Dim rsMicroseguro As ADODB.Recordset
    Dim oPoliza As COMDCredito.DCOMPoliza
    Dim sCtaCod As String
    Dim sTelefono As String
    
    Set oPoliza = New COMDCredito.DCOMPoliza
    
    On Error GoTo ErrorTrama

    If sTipoD = "AF" Then
        Set rsMicroseguro = oPoliza.GenerarTramaMicroseguroAF(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin, Trim(Right(Me.cmbMicroseguros.Text, 2)))
    ElseIf sTipoD = "CA" Then
        Set rsMicroseguro = oPoliza.GenerarTramaMicroseguroCA(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin, Trim(Right(Me.cmbMicroseguros.Text, 2)))
    End If
    
    If rsMicroseguro.RecordCount > 0 Then
        nMonto = ObtenerMonto(rsMicroseguro)
        
        'Abre el archivo excel
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        
        If sTipoD = "AF" Then
            lsArchivo = "TramaMicrosegurosAF"
            lsNomHoja = "TramaMicrosegurosAF"
        ElseIf sTipoD = "CA" Then
            lsArchivo = "TramaMicrosegurosCA"
            lsNomHoja = "TramaMicrosegurosCA"
        End If
        
        lsArchivo1 = "\spooler\" & lsArchivo & "_" & gsCodUser & "_" & Format(gdFecha, "yyyymmdd") & "_" & Format$(Time(), "HHMMSS") & ".xls"
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
            
            xlHoja1.Cells(3, 1) = Me.txtCodPatrocinador.Text
            xlHoja1.Cells(3, 2) = Me.txtNumRamo.Text
            xlHoja1.Cells(3, 3) = Me.txtCodProducto.Text
            xlHoja1.Cells(3, 4) = Me.txtNumPoliza.Text
            xlHoja1.Cells(3, 5) = sTipoD
            xlHoja1.Cells(3, 6) = Format(gdFecSis, "dd/mm/yyyy")
            xlHoja1.Cells(3, 7) = Format(dfechaini, "dd/mm/yyyy")
            xlHoja1.Cells(3, 8) = Format(dfechafin, "dd/mm/yyyy")
            If sTipoD = "AF" Then
                nNumMov = ObtenerNumMov(rsMicroseguro, 1)
                nNumRep = ObtenerNumMov(rsMicroseguro, 2) + ObtenerNumMov(rsMicroseguro, 3)
                xlHoja1.Cells(3, 9) = nNumMov
                xlHoja1.Cells(3, 10) = (nNumMov * 2) + nNumRep
                xlHoja1.Cells(3, 11) = CStr(nMonto)
            ElseIf sTipoD = "CA" Then
                nNumMov = rsMicroseguro.RecordCount
                xlHoja1.Cells(3, 9) = nNumMov
                xlHoja1.Cells(3, 10) = nNumMov + 1
                xlHoja1.Cells(3, 11) = CStr(nMonto)
            End If
            xlHoja1.Cells(3, 12) = sMonedaTXT

            rsMicroseguro.MoveFirst
            If Not (rsMicroseguro.EOF And rsMicroseguro.BOF) Then
                If sTipoD = "AF" Then
                    nNumLineas = 7
                    sCtaCod = ""
                    For i = 0 To rsMicroseguro.RecordCount - 1
                        If Trim(sCtaCod) <> Trim(rsMicroseguro!cCtaCod) Then
                            'DECLARACION
                            xlHoja1.Cells(nNumLineas, 1) = "0"
                            xlHoja1.Cells(nNumLineas, 2) = Format(gdFecSis, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 3) = rsMicroseguro!CodSucVentPat
                            xlHoja1.Cells(nNumLineas, 4) = rsMicroseguro!CodVendPat
                            xlHoja1.Cells(nNumLineas, 5) = String(10, " ")
                            xlHoja1.Cells(nNumLineas, 6) = Format(rsMicroseguro!FechaInicio, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 7) = Format(rsMicroseguro!fechafin, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 8) = String(3, "0")
                            xlHoja1.Cells(nNumLineas, 9) = Format(txtCodPlan.Text, "00000")
                            xlHoja1.Cells(nNumLineas, 10) = rsMicroseguro!Moneda
                            xlHoja1.Cells(nNumLineas, 11) = Format("01/01/1900", "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 12) = String(13, "0")
                            xlHoja1.Cells(nNumLineas, 13) = CDbl(rsMicroseguro!PrimaMensual)
                            xlHoja1.Cells(nNumLineas, 14) = CDbl(rsMicroseguro!PrimaMensual)
                            xlHoja1.Cells(nNumLineas, 15) = String(3, "0")
                            xlHoja1.Cells(nNumLineas, 16) = rsMicroseguro!NroCert
                            xlHoja1.Cells(nNumLineas, 17) = "AF"
                            xlHoja1.Cells(nNumLineas, 18) = String(11, " ")
                            
                            'CONTRATANTE
                            xlHoja1.Cells(nNumLineas, 19) = "1"
                            xlHoja1.Cells(nNumLineas, 20) = Format(rsMicroseguro!TpoPersContr, "0")
                            xlHoja1.Cells(nNumLineas, 21) = Format(rsMicroseguro!TpoDocIDContr, "0")
                            xlHoja1.Cells(nNumLineas, 22) = rsMicroseguro!NroDocIDContr
                            xlHoja1.Cells(nNumLineas, 22) = String(20, " ")
                            xlHoja1.Cells(nNumLineas, 24) = String(20, " ")
                            xlHoja1.Cells(nNumLineas, 25) = String(30, " ")
                            xlHoja1.Cells(nNumLineas, 26) = String(8, " ")
                            xlHoja1.Cells(nNumLineas, 27) = rsMicroseguro!SexoContr
                            xlHoja1.Cells(nNumLineas, 28) = rsMicroseguro!DirContr
                            xlHoja1.Cells(nNumLineas, 29) = Format(rsMicroseguro!UbiDepContr, "00")
                            xlHoja1.Cells(nNumLineas, 30) = Format(rsMicroseguro!UbiProvContr, "00")
                            xlHoja1.Cells(nNumLineas, 31) = Format(rsMicroseguro!UbiDistrContr, "00")
                            xlHoja1.Cells(nNumLineas, 32) = String(12, " ")
                            xlHoja1.Cells(nNumLineas, 33) = String(12, " ")
                            xlHoja1.Cells(nNumLineas, 34) = String(1, " ")
                            xlHoja1.Cells(nNumLineas, 35) = rsMicroseguro!RazonSocContr
                            xlHoja1.Cells(nNumLineas, 36) = String(30, " ")
                            xlHoja1.Cells(nNumLineas, 37) = String(20, " ")
                            sTelefono = rsMicroseguro!TelfContr
                            sTelefono = QuitarCaracter(QuitarCaracter(QuitarCaracter(sTelefono, "-"), ")"), "(")
                            xlHoja1.Cells(nNumLineas, 38) = sTelefono
                            xlHoja1.Cells(nNumLineas, 39) = rsMicroseguro!NroCertContr
                            
                            'ASEGURADO
                            xlHoja1.Cells(nNumLineas, 40) = "2"
                            xlHoja1.Cells(nNumLineas, 41) = rsMicroseguro!TpoPersAseg
                            xlHoja1.Cells(nNumLineas, 42) = Format(rsMicroseguro!TpoDocIDAseg, "0")
                            xlHoja1.Cells(nNumLineas, 43) = rsMicroseguro!NroDocIDAseg
                            xlHoja1.Cells(nNumLineas, 44) = rsMicroseguro!ApePatAseg
                            xlHoja1.Cells(nNumLineas, 45) = rsMicroseguro!ApeMatAseg
                            xlHoja1.Cells(nNumLineas, 46) = rsMicroseguro!NombreAseg
                            xlHoja1.Cells(nNumLineas, 47) = Format(rsMicroseguro!FecNacAseg, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 48) = rsMicroseguro!SexoAseg
                            xlHoja1.Cells(nNumLineas, 49) = rsMicroseguro!DirAseg
                            xlHoja1.Cells(nNumLineas, 50) = Format(rsMicroseguro!UbiDepAseg, "00")
                            xlHoja1.Cells(nNumLineas, 51) = Format(rsMicroseguro!UbiProvAseg, "00")
                            xlHoja1.Cells(nNumLineas, 52) = Format(rsMicroseguro!UbiDistrAseg, "00")
                            xlHoja1.Cells(nNumLineas, 53) = rsMicroseguro!TelfAseg
                            xlHoja1.Cells(nNumLineas, 54) = rsMicroseguro!CelAseg
                            xlHoja1.Cells(nNumLineas, 55) = rsMicroseguro!EstCivAseg
                            xlHoja1.Cells(nNumLineas, 56) = String(1, "0")
                            xlHoja1.Cells(nNumLineas, 57) = rsMicroseguro!NroCertAseg
                        End If
                        If Trim(rsMicroseguro!NroDocIDBenf) <> "" Then
                            'BENEFICIARIO
                            xlHoja1.Cells(nNumLineas, 58) = "3"
                            xlHoja1.Cells(nNumLineas, 59) = rsMicroseguro!NroDocIDBenf
                            xlHoja1.Cells(nNumLineas, 60) = rsMicroseguro!ApePatBenf
                            xlHoja1.Cells(nNumLineas, 61) = rsMicroseguro!ApeMatBenf
                            xlHoja1.Cells(nNumLineas, 62) = rsMicroseguro!NombreBenf
                            xlHoja1.Cells(nNumLineas, 63) = Format(rsMicroseguro!FecNacBenf, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 64) = rsMicroseguro!ParentBenf
                            xlHoja1.Cells(nNumLineas, 65) = (CInt(rsMicroseguro!PorcBenf) / 100) & "%"
                            xlHoja1.Cells(nNumLineas, 66) = rsMicroseguro!NroCertBenf
                        End If
                        If Trim(sCtaCod) <> Trim(rsMicroseguro!cCtaCod) Then
                            If Trim(rsMicroseguro!NroDocIDCony) <> "" Then
                                'CONYUGE
                                xlHoja1.Cells(nNumLineas, 67) = "4"
                                xlHoja1.Cells(nNumLineas, 68) = rsMicroseguro!NroDocIDCony
                                xlHoja1.Cells(nNumLineas, 69) = rsMicroseguro!ApePatCony
                                xlHoja1.Cells(nNumLineas, 70) = rsMicroseguro!ApeMatCony
                                xlHoja1.Cells(nNumLineas, 71) = rsMicroseguro!NombreCony
                                xlHoja1.Cells(nNumLineas, 72) = Format(rsMicroseguro!FecNacCony, "dd/mm/yyyy")
                                xlHoja1.Cells(nNumLineas, 73) = rsMicroseguro!NroCertCony
                            End If
                        End If
                        xlHoja1.Range(xlHoja1.Cells(nNumLineas, 1), xlHoja1.Cells(nNumLineas, 73)).Borders.LineStyle = 1
                        nNumLineas = nNumLineas + 1
                        sCtaCod = rsMicroseguro!cCtaCod
                        rsMicroseguro.MoveNext
                    Next i
                ElseIf sTipoD = "CA" Then
                    nNumLineas = 7
                    'DEPOSITO
                    xlHoja1.Cells(nNumLineas, 1) = "D"
                    xlHoja1.Cells(nNumLineas, 2) = CStr(nMonto)
                    xlHoja1.Cells(nNumLineas, 3) = sTipoPago
                    xlHoja1.Cells(nNumLineas, 4) = txtCta.Text
                    xlHoja1.Cells(nNumLineas, 5) = Me.txtMesDepostito.Text
                    xlHoja1.Cells(nNumLineas, 6) = txtBanco.Text
                    xlHoja1.Cells(nNumLineas, 7) = txtNroDocumento.Text
                    xlHoja1.Cells(nNumLineas, 8) = txtObservaciones.Text
                    
                    nNumLineas = 11
                    For i = 0 To rsMicroseguro.RecordCount - 1
                        'DETALLE RECAUDACION
                        xlHoja1.Cells(nNumLineas, 1) = "C"
                        xlHoja1.Cells(nNumLineas, 2) = rsMicroseguro!CtaCli
                        xlHoja1.Cells(nNumLineas, 3) = rsMicroseguro!NombreCli
                        xlHoja1.Cells(nNumLineas, 4) = CDbl(rsMicroseguro!PrimaMensual)
                        xlHoja1.Cells(nNumLineas, 5) = CDbl(rsMicroseguro!PrimaMensualB)
                        xlHoja1.Cells(nNumLineas, 6) = rsMicroseguro!NroCert
                        xlHoja1.Cells(nNumLineas, 7) = rsMicroseguro!Observaciones
                        xlHoja1.Range(xlHoja1.Cells(nNumLineas, 1), xlHoja1.Cells(nNumLineas, 7)).Borders.LineStyle = 1
                        nNumLineas = nNumLineas + 1
                        rsMicroseguro.MoveNext
                    Next i
                    
                End If
            End If

            nCerrar = 0
            
         xlHoja1.SaveAs App.path & lsArchivo1
         psArchivoAGrabarC = App.path & lsArchivo1
         xlsAplicacion.Visible = True
         xlsAplicacion.Windows(1).Visible = True
         Set xlsAplicacion = Nothing
         Set xlsLibro = Nothing
         Set xlHoja1 = Nothing
        MsgBox "Trama de Microseguros Generada Satisfactoriamente. RUTA: " & psArchivoAGrabarC, vbInformation, "Aviso"
    Else
        MsgBox "No hay Datos", vbInformation, "Aviso"
    End If
    
    Exit Sub
ErrorTrama:
    MsgBox Err.Description, vbInformation, "Aviso de Error"
End Sub
Private Sub MultiriesgoEXCEL()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nCerrar As Integer
    Dim i As Integer
    Dim nNumLineas As Integer
    Dim nNumRep As Integer
    Dim nNumMov As Integer
    Dim rsMultiriesgo As ADODB.Recordset
    Dim oPoliza As COMDCredito.DCOMPoliza
    Dim sCtaCod As String
    Dim sTelefono As String
    Dim nContMat As Integer
    
    Set oPoliza = New COMDCredito.DCOMPoliza
    
    On Error GoTo ErrorTrama

    If sTipoD = "AF" Then
        Set rsMultiriesgo = oPoliza.GenerarTramaMultiriesgoAF(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin)
    ElseIf sTipoD = "CA" Then
        Set rsMultiriesgo = oPoliza.GenerarTramaMultiriesgoCA(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin)
    End If
    If rsMultiriesgo.RecordCount > 0 Then
        nMonto = ObtenerMonto(rsMultiriesgo)
        
        'Abre el archivo excel
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        
        If sTipoD = "AF" Then
            lsArchivo = "TramaMultiriesgoAF"
            lsNomHoja = "TramaMultiriesgoAF"
        ElseIf sTipoD = "CA" Then
            lsArchivo = "TramaMultiriesgoCA"
            lsNomHoja = "TramaMultiriesgoCA"
        End If
        
        lsArchivo1 = "\spooler\" & lsArchivo & "_" & gsCodUser & "_" & Format(gdFecha, "yyyymmdd") & "_" & Format$(Time(), "HHMMSS") & ".xls"
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
            
            xlHoja1.Cells(3, 1) = Me.txtCodPatrocinador.Text
            xlHoja1.Cells(3, 2) = Me.txtNumRamo.Text
            xlHoja1.Cells(3, 3) = Me.txtCodProducto.Text
            xlHoja1.Cells(3, 4) = Me.txtNumPoliza.Text
            xlHoja1.Cells(3, 5) = sTipoD
            xlHoja1.Cells(3, 6) = Format(gdFecSis, "dd/mm/yyyy")
            xlHoja1.Cells(3, 7) = Format(dfechaini, "dd/mm/yyyy")
            xlHoja1.Cells(3, 8) = Format(dfechafin, "dd/mm/yyyy")
            If sTipoD = "AF" Then
                nNumMov = ObtenerNumMov(rsMultiriesgo, 1)
                nNumRep = ObtenerNumMov(rsMultiriesgo, 4)
                xlHoja1.Cells(3, 9) = nNumMov
                xlHoja1.Cells(3, 10) = (nNumMov * 4) + nNumRep
                xlHoja1.Cells(3, 11) = CStr(nMonto)
            ElseIf sTipoD = "CA" Then
                nNumMov = rsMultiriesgo.RecordCount
                xlHoja1.Cells(3, 9) = nNumMov
                xlHoja1.Cells(3, 10) = nNumMov + 1
                xlHoja1.Cells(3, 11) = CStr(nMonto)
            End If
            xlHoja1.Cells(3, 12) = sMonedaTXT

            rsMultiriesgo.MoveFirst
            If Not (rsMultiriesgo.EOF And rsMultiriesgo.BOF) Then
                If sTipoD = "AF" Then
                    nNumLineas = 7
                    sCtaCod = ""
                    For i = 0 To rsMultiriesgo.RecordCount - 1
                        If Trim(sCtaCod) <> Trim(rsMultiriesgo!cCtaCod) Then
                            nContMat = 1
                            'DECLARACION
                            xlHoja1.Cells(nNumLineas, 1) = "0"
                            xlHoja1.Cells(nNumLineas, 2) = rsMultiriesgo!NumCert
                            xlHoja1.Cells(nNumLineas, 3) = Format(rsMultiriesgo!FecVigInicial, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 4) = Format(rsMultiriesgo!FecVigFinal, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 5) = CDbl(rsMultiriesgo!PrimaMensual)
                            xlHoja1.Cells(nNumLineas, 6) = rsMultiriesgo!CodPlan
                                         
                            'DECLARACION JURADA
                            xlHoja1.Cells(nNumLineas, 7) = "1"
                            xlHoja1.Cells(nNumLineas, 8) = rsMultiriesgo!TpoDJ
                            xlHoja1.Cells(nNumLineas, 9) = rsMultiriesgo!NroDJ
                            xlHoja1.Cells(nNumLineas, 10) = rsMultiriesgo!TpoGarDJ
                            xlHoja1.Cells(nNumLineas, 11) = Format(rsMultiriesgo!FechaDJ, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 12) = rsMultiriesgo!AntNroDJ
                            xlHoja1.Cells(nNumLineas, 13) = rsMultiriesgo!NomCliDJ
                            xlHoja1.Cells(nNumLineas, 14) = rsMultiriesgo!DniCliDJ
                            xlHoja1.Cells(nNumLineas, 15) = rsMultiriesgo!NomConyDJ
                            xlHoja1.Cells(nNumLineas, 16) = rsMultiriesgo!DniConyDJ
                            xlHoja1.Cells(nNumLineas, 17) = rsMultiriesgo!DirDJ
                            xlHoja1.Cells(nNumLineas, 18) = rsMultiriesgo!MonedaDJ
                            xlHoja1.Cells(nNumLineas, 19) = rsMultiriesgo!MontoDJ
                            xlHoja1.Cells(nNumLineas, 20) = rsMultiriesgo!NroCertDJ
                            
                            'MATERIAS ASEGUTRADAS
                            xlHoja1.Cells(nNumLineas, 21) = "2"
                            xlHoja1.Cells(nNumLineas, 22) = rsMultiriesgo!NroDJMatAseg
                            xlHoja1.Cells(nNumLineas, 23) = rsMultiriesgo!TpoGarMatAseg
                            xlHoja1.Cells(nNumLineas, 24) = nContMat 'rsMultiriesgo!ContMatAseg
                            xlHoja1.Cells(nNumLineas, 25) = rsMultiriesgo!DirMatAseg
                            xlHoja1.Cells(nNumLineas, 26) = Format(rsMultiriesgo!UbiDepMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 27) = Format(rsMultiriesgo!UbiProvMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 28) = Format(rsMultiriesgo!UbiDistMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 29) = rsMultiriesgo!AnoConsMatAseg
                            xlHoja1.Cells(nNumLineas, 30) = rsMultiriesgo!NroPisMatAseg
                            xlHoja1.Cells(nNumLineas, 31) = rsMultiriesgo!UsoMatAseg
                            xlHoja1.Cells(nNumLineas, 32) = rsMultiriesgo!MatMatAseg
                            xlHoja1.Cells(nNumLineas, 33) = rsMultiriesgo!TpoMatAseg
                            xlHoja1.Cells(nNumLineas, 34) = rsMultiriesgo!DescMatAseg
                            xlHoja1.Cells(nNumLineas, 35) = rsMultiriesgo!NroCertMatAseg
                        
                            'ASEGURADO
                            xlHoja1.Cells(nNumLineas, 36) = "4"
                            xlHoja1.Cells(nNumLineas, 37) = Format(rsMultiriesgo!TpoDocIDAseg, "0")
                            xlHoja1.Cells(nNumLineas, 38) = rsMultiriesgo!NroDocIDAseg
                            xlHoja1.Cells(nNumLineas, 39) = rsMultiriesgo!ApePatAseg
                            xlHoja1.Cells(nNumLineas, 40) = rsMultiriesgo!ApeMatAseg
                            xlHoja1.Cells(nNumLineas, 41) = rsMultiriesgo!NombreAseg
                            xlHoja1.Cells(nNumLineas, 42) = Format(rsMultiriesgo!FecNacAseg, "dd/mm/yyyy")
                            xlHoja1.Cells(nNumLineas, 43) = rsMultiriesgo!SexoAseg
                            xlHoja1.Cells(nNumLineas, 44) = rsMultiriesgo!DirAseg
                            xlHoja1.Cells(nNumLineas, 45) = Format(rsMultiriesgo!UbiDepAseg, "00")
                            xlHoja1.Cells(nNumLineas, 46) = Format(rsMultiriesgo!UbiProvAseg, "00")
                            xlHoja1.Cells(nNumLineas, 47) = Format(rsMultiriesgo!UbiDistrAseg, "00")
                            xlHoja1.Cells(nNumLineas, 48) = rsMultiriesgo!TelfAseg
                            xlHoja1.Cells(nNumLineas, 49) = rsMultiriesgo!EstCivAseg
                            xlHoja1.Cells(nNumLineas, 50) = rsMultiriesgo!NroCertAseg
                        End If
                        If Trim(sCtaCod) = Trim(rsMultiriesgo!cCtaCod) Then
                            'MATERIAS ASEGUTRADAS
                            xlHoja1.Cells(nNumLineas, 21) = "2"
                            xlHoja1.Cells(nNumLineas, 22) = rsMultiriesgo!NroDJMatAseg
                            xlHoja1.Cells(nNumLineas, 23) = rsMultiriesgo!TpoGarMatAseg
                            xlHoja1.Cells(nNumLineas, 24) = nContMat 'rsMultiriesgo!ContMatAseg
                            xlHoja1.Cells(nNumLineas, 25) = rsMultiriesgo!DirMatAseg
                            xlHoja1.Cells(nNumLineas, 26) = Format(rsMultiriesgo!UbiDepMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 27) = Format(rsMultiriesgo!UbiProvMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 28) = Format(rsMultiriesgo!UbiDistMatAseg, "00")
                            xlHoja1.Cells(nNumLineas, 29) = rsMultiriesgo!AnoConsMatAseg
                            xlHoja1.Cells(nNumLineas, 30) = rsMultiriesgo!NroPisMatAseg
                            xlHoja1.Cells(nNumLineas, 31) = rsMultiriesgo!UsoMatAseg
                            xlHoja1.Cells(nNumLineas, 32) = rsMultiriesgo!MatMatAseg
                            xlHoja1.Cells(nNumLineas, 33) = rsMultiriesgo!TpoMatAseg
                            xlHoja1.Cells(nNumLineas, 34) = rsMultiriesgo!DescMatAseg
                            xlHoja1.Cells(nNumLineas, 35) = rsMultiriesgo!NroCertMatAseg
                        End If
                       
                        xlHoja1.Range(xlHoja1.Cells(nNumLineas, 1), xlHoja1.Cells(nNumLineas, 50)).Borders.LineStyle = 1
                        nNumLineas = nNumLineas + 1
                        sCtaCod = rsMultiriesgo!cCtaCod
                        nContMat = nContMat + 1
                        rsMultiriesgo.MoveNext
                    Next i
                ElseIf sTipoD = "CA" Then
                    nNumLineas = 7
                    'DEPOSITO
                    xlHoja1.Cells(nNumLineas, 1) = "D"
                    xlHoja1.Cells(nNumLineas, 2) = CStr(nMonto)
                    xlHoja1.Cells(nNumLineas, 3) = sTipoPago
                    xlHoja1.Cells(nNumLineas, 4) = txtCta.Text
                    xlHoja1.Cells(nNumLineas, 5) = Me.txtMesDepostito.Text
                    xlHoja1.Cells(nNumLineas, 6) = txtBanco.Text
                    xlHoja1.Cells(nNumLineas, 7) = txtNroDocumento.Text
                    xlHoja1.Cells(nNumLineas, 8) = txtObservaciones.Text
                    
                    nNumLineas = 11
                    For i = 0 To rsMultiriesgo.RecordCount - 1
                        'DETALLE RECAUDACION
                        xlHoja1.Cells(nNumLineas, 1) = "C"
                        xlHoja1.Cells(nNumLineas, 2) = rsMultiriesgo!CtaCli
                        xlHoja1.Cells(nNumLineas, 3) = rsMultiriesgo!NombreCli
                        xlHoja1.Cells(nNumLineas, 4) = CDbl(rsMultiriesgo!PrimaMensual)
                        xlHoja1.Cells(nNumLineas, 5) = CDbl(rsMultiriesgo!PrimaMensualB)
                        xlHoja1.Cells(nNumLineas, 6) = rsMultiriesgo!NroCert
                        xlHoja1.Cells(nNumLineas, 7) = rsMultiriesgo!Observaciones
                        xlHoja1.Range(xlHoja1.Cells(nNumLineas, 1), xlHoja1.Cells(nNumLineas, 7)).Borders.LineStyle = 1
                        nNumLineas = nNumLineas + 1
                        rsMultiriesgo.MoveNext
                    Next i
                End If
            End If
         xlHoja1.SaveAs App.path & lsArchivo1
         psArchivoAGrabarC = App.path & lsArchivo1
         xlsAplicacion.Visible = True
         xlsAplicacion.Windows(1).Visible = True
         Set xlsAplicacion = Nothing
         Set xlsLibro = Nothing
         Set xlHoja1 = Nothing
        MsgBox "Trama de Multiriesgo Generada Satisfactoriamente. RUTA: " & psArchivoAGrabarC, vbInformation, "Aviso"
    Else
        MsgBox "No hay Datos", vbInformation, "Aviso"
    End If
    Exit Sub
ErrorTrama:
    MsgBox Err.Description, vbInformation, "Aviso de Error"
End Sub


Private Sub cmbMoneda_Click()
    If Trim(Right(cmbMoneda.Text, 2)) = "1" Then
        sMonedaTXT = "S"
    ElseIf Trim(Right(cmbMoneda.Text, 2)) = "2" Then
        sMonedaTXT = "D"
    End If
End Sub

Private Sub cmbTipoDeclaracion_Click()
sTipoD = Left(Right(Me.cmbTipoDeclaracion.Text, 3), 2)
    If sTipoD = "CA" Then
        Me.txtCodPlan.Visible = False
        Me.Label15.Visible = False
        Me.Height = 7600
    Else
        If sTipoTrama = "1" Then
            Me.txtCodPlan.Visible = True
            Me.Label15.Visible = True
        End If
        Me.Height = 4400
    End If
End Sub



Private Sub cmbTipoPago_Click()
sTipoPago = Trim(Right(Me.cmbTipoPago.Text, 3))
End Sub

Private Sub cmdGenerar_Click()
    If validaDatos Then
        dfechaini = "01/" & Me.txtMesActual
        dfechafin = DateDiff("d", CDate(dfechaini), DateAdd("M", 1, CDate(dfechaini))) & "/" & Me.txtMesActual
        If Me.opteExportarTXT.value = True Then
            If sTipoTrama = "1" Then
                Call Microseguros
            ElseIf sTipoTrama = "2" Then
                Call Multiriesgos
            End If
         ElseIf Me.optExportarExcel.value = True Then
            If sTipoTrama = "1" Then
                Call MicrosegurosEXCEL
            ElseIf sTipoTrama = "2" Then
                Call MultiriesgoEXCEL
            End If
         End If
    End If
End Sub

Private Sub cmdTipo_Click()
sTipoTrama = Trim(Right(Me.cmdTipo.Text, 3))
    If sTipoD = "CA" Then
        Me.txtCodPlan.Visible = False
        Me.Label15.Visible = False
        Me.Height = 7600
    Else
        If sTipoTrama = "1" Then
            Me.txtCodPlan.Visible = True
            Me.Label15.Visible = True
        Else
            Me.txtCodPlan.Visible = False
            Me.Label15.Visible = False
        End If
    Me.Height = 4400
    End If
    If sTipoTrama = "1" Then
        Me.frmMicroseguro.Visible = True
        cmbMicroseguros.ListIndex = IndiceListaCombo(cmbMicroseguros, 0)
    Else
        Me.frmMicroseguro.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    bValidaFec = False
    Me.txtMesActual.Text = Mid(CStr(DateAdd("M", -1, gdFecSis)), 4, 7)
    Call ValidaFecha
    Call Cargar_Objetos_Controles
    Me.Height = 4400
    Me.txtMesDepostito.Text = gdFecSis
    Me.frmMicroseguro.Visible = False
End Sub

Private Sub Cargar_Objetos_Controles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rsConstante As ADODB.Recordset
Dim rsAgencias As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCreditos

'Carga Monedas
Set oCons = New COMDConstantes.DCOMConstantes
Set rsConstante = oCons.RecuperaConstantes(gMoneda)
Call Llenar_Combo_con_Recordset(rsConstante, cmbMoneda)

'Carga Agencias
Set oCred = New COMDCredito.DCOMCreditos
Set rsAgencias = oCred.ObtieneAgenciasAdmCred
Call CargaAgencias(rsAgencias)

'Cargar Tipo Microseguro
Set rsConstante = oCons.RecuperaConstantes(9992)
Call Llenar_Combo_con_Recordset(rsConstante, cmbMicroseguros)

Set oCred = Nothing
Set rsConstante = Nothing
Set rsAgencias = Nothing
End Sub
Sub CargaAgencias(ByVal pRs As ADODB.Recordset)
On Error GoTo ErrHandler
        Do Until pRs.EOF
            Me.cmbAgencia.AddItem Left(pRs!cAgeDescripcion, 40) + Space(50) + pRs!cAgeCod
            pRs.MoveNext
        Loop
Exit Sub
ErrHandler:
    MsgBox "Error al cargar CargaAgenciasAdmCred", vbInformation, "AVISO"
End Sub

Private Sub Microseguros()
    Dim CadTemp As String
    Dim sCad As String
    Dim ArcSal As Integer
    Dim rsMicroseguro As ADODB.Recordset
    Dim nNumMov As Integer
    Dim nNumRep As Integer
    Dim oPoliza As COMDCredito.DCOMPoliza
    Set oPoliza = New COMDCredito.DCOMPoliza
    On Error GoTo ErrorTrama
    
    If sTipoD = "AF" Then
        Set rsMicroseguro = oPoliza.GenerarTramaMicroseguroAF(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin, Trim(Right(Me.cmbMicroseguros.Text, 2)))
    ElseIf sTipoD = "CA" Then
        Set rsMicroseguro = oPoliza.GenerarTramaMicroseguroCA(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin, Trim(Right(Me.cmbMicroseguros.Text, 2)))
    End If

    If rsMicroseguro.RecordCount > 0 Then
    
        nMonto = ObtenerMonto(rsMicroseguro)
        
        psArchivoAGrabarC = App.path & "\SPOOLER\CMAC_MAYNAS_AP_" & sTipoD & _
                        "_" & Format(gdFecSis, "ddMMyy") & "" & _
                        Format(Time, "hhmmss") & "_" & IIf(Trim(Right(Me.cmbMicroseguros.Text, 2)) = "1", "S250", "S150") & ".txt"
        
        CadTemp = Format(Trim(Me.txtCodPatrocinador.Text), "000000") & _
                Format(Trim(Me.txtNumRamo.Text), "000") & _
                Format(Trim(Me.txtCodProducto.Text), "000000") & _
                Format(Trim(Me.txtNumPoliza.Text), "000000000000") & _
                sTipoD & _
                Format(gdFecSis, "yyyymmdd") & _
                Format(dfechaini, "yyyymmdd") & _
                Format(dfechafin, "yyyymmdd")
        If sTipoD = "AF" Then
            nNumMov = ObtenerNumMov(rsMicroseguro, 1)
            nNumRep = ObtenerNumMov(rsMicroseguro, 2) + ObtenerNumMov(rsMicroseguro, 3)
            CadTemp = CadTemp & Format(nNumMov, "00000") & _
            Format((nNumMov * 2) + nNumRep, "00000") & _
            Format(nMonto * 100, "000000000")
        ElseIf sTipoD = "CA" Then
            nNumMov = rsMicroseguro.RecordCount
            CadTemp = CadTemp & Format(nNumMov, "000000") & _
            Format(nNumMov + 1, "000000") & _
            Format(nMonto * 100, "0000000000000")
        End If
        CadTemp = CadTemp & sMonedaTXT
                
        ArcSal = FreeFile
        sCad = ""
        Open psArchivoAGrabarC For Output As ArcSal
            If CadTemp <> "" Then
                Print #1, sCad; CadTemp
                If sTipoD = "AF" Then
                    Call EstructuraAfiliaciones(rsMicroseguro, CInt(sTipoTrama))
                ElseIf sTipoD = "CA" Then
                    Call EstructuraCargos(rsMicroseguro, CInt(sTipoTrama))
                End If
            End If
        Close ArcSal
    MsgBox "Trama de Microseguros Generada Satisfactoriamente. RUTA: " & psArchivoAGrabarC, vbInformation, "Aviso"
    Else
        MsgBox "No hay Datos", vbInformation, "Aviso"
    End If
    Exit Sub
ErrorTrama:
    MsgBox Err.Description, vbInformation, "Aviso de Error"
End Sub

Private Sub Multiriesgos()
Dim CadTemp As String
    Dim sCad As String
    Dim ArcSal As Integer
    Dim rsMultiriesgo As ADODB.Recordset
    Dim nMonto As Double
    Dim nNumMov As Integer
    Dim nNumRep As Integer
    Dim oPoliza As COMDCredito.DCOMPoliza
    Set oPoliza = New COMDCredito.DCOMPoliza
    
    On Error GoTo ErrorTrama
    If sTipoD = "AF" Then
        Set rsMultiriesgo = oPoliza.GenerarTramaMultiriesgoAF(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin)
    ElseIf sTipoD = "CA" Then
        Set rsMultiriesgo = oPoliza.GenerarTramaMultiriesgoCA(Trim(Right(Me.cmbAgencia.Text, 4)), Trim(Right(Me.cmbMoneda.Text, 2)), dfechaini, dfechafin)
    End If
    If rsMultiriesgo.RecordCount > 0 Then
    nMonto = ObtenerMonto(rsMultiriesgo)
    psArchivoAGrabarC = App.path & "\SPOOLER\CMACMAYNAS_MULTI_" & sTipoD & _
                    "_" & Format(gdFecSis, "ddMMyy") & "" & _
                    Format(Time, "hhmmss") & ".txt"
    
    CadTemp = Format(Trim(Me.txtCodPatrocinador.Text), "000000") & _
            Format(Trim(Me.txtNumRamo.Text), "000") & _
            Format(Trim(Me.txtCodProducto.Text), "000000") & _
            Format(Trim(Me.txtNumPoliza.Text), "000000000000") & _
            sTipoD & _
            Format(gdFecSis, "yyyymmdd") & _
            Format(dfechaini, "yyyymmdd") & _
            Format(dfechafin, "yyyymmdd")
    If sTipoD = "AF" Then
        nNumMov = ObtenerNumMov(rsMultiriesgo, 1)
        nNumRep = ObtenerNumMov(rsMultiriesgo, 4)
        CadTemp = CadTemp & Format(nNumMov, "00000") & _
        Format((nNumMov * 4) + nNumRep, "00000") & _
        Format(nMonto * 100, "000000000")
    ElseIf sTipoD = "CA" Then
        nNumMov = rsMultiriesgo.RecordCount
        CadTemp = CadTemp & Format(nNumMov, "000000") & _
        Format(nNumMov + 1, "000000") & _
        Format(nMonto * 100, "0000000000000")
    End If
    CadTemp = CadTemp & sMonedaTXT
            
    ArcSal = FreeFile
    sCad = ""
    Open psArchivoAGrabarC For Output As ArcSal
        If CadTemp <> "" Then
           Print #1, sCad; CadTemp
        If sTipoD = "AF" Then
            Call EstructuraAfiliaciones(rsMultiriesgo, CInt(sTipoTrama))
        ElseIf sTipoD = "CA" Then
            Call EstructuraCargos(rsMultiriesgo, CInt(sTipoTrama))
        End If
        
        End If
    Close ArcSal
    MsgBox "Trama de Multiriesgos Generada Satisfactoriamente. RUTA: " & psArchivoAGrabarC, vbInformation, "Aviso"
    Else
        MsgBox "No hay Datos", vbInformation, "Aviso"
    End If
    Exit Sub
ErrorTrama:
    MsgBox Err.Description, vbInformation, "Aviso de Error"
End Sub

Private Sub EstructuraAfiliaciones(ByVal pRs As ADODB.Recordset, ByVal pnTipo As Integer)
Dim sCad0 As String
Dim sCad1 As String
Dim sCad2 As String
Dim sCad3 As String
Dim sCad4 As String
Dim nConteo As Integer
Dim sTelefono As String
Dim sCtaCod As String
Dim nContMat As Integer
pRs.MoveFirst
If pnTipo = "1" Then
    If Not (pRs.EOF And pRs.BOF) Then
        sCtaCod = ""
        For nConteo = 1 To pRs.RecordCount
            If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
             'DECLARACION
             sCad0 = "0" & Format(gdFecSis, "yyyymmdd")
             sCad0 = sCad0 & Format(pRs!CodSucVentPat, "0000")
             sCad0 = sCad0 & IIf(Len(pRs!CodVendPat) > 15, Mid(pRs!CodVendPat, 1, 15), _
                   IIf(Len(pRs!CodVendPat) < 15, pRs!CodVendPat & String(Abs(15 - Len(pRs!CodVendPat)), " "), _
                   pRs!CodVendPat))
             sCad0 = sCad0 & String(10, "0")
             sCad0 = sCad0 & Format(pRs!FechaInicio, "yyyymmdd")
             sCad0 = sCad0 & Format(pRs!fechafin, "yyyymmdd")
             sCad0 = sCad0 & String(3, "0")
             sCad0 = sCad0 & Format(txtCodPlan.Text, "00000")
             sCad0 = sCad0 & pRs!Moneda
             sCad0 = sCad0 & Format("01/01/1900", "yyyymmdd")
             sCad0 = sCad0 & String(13, "0")
             sCad0 = sCad0 & Format((CDbl(pRs!PrimaMensual)) * 100, "0000000000000") ' Format(CDbl(50) * 100, "0000000000000") 'String(13, "0")
             sCad0 = sCad0 & Format(CDbl(pRs!PrimaMensual) * 100, "0000000000000") 'Format(CDbl(4.17) * 100, "0000000000000")
             sCad0 = sCad0 & String(3, "0")
             sCad0 = sCad0 & Format(QuitarCaracter(pRs!NroCert, "-"), "0000000000")
             sCad0 = sCad0 & "AF"
             sCad0 = sCad0 & String(11, " ")
             Print #1, ""; sCad0
             
            'CONTRATANTE
'             sCad1 = "1" & Format(pRs!TpoPersContr, "0")
'             sCad1 = sCad1 & Format(pRs!TpoDocIDContr, "0")
'             sCad1 = sCad1 & IIf(Len(pRs!NroDocIDContr) > 15, Mid(pRs!NroDocIDContr, 1, 15), _
'                   IIf(Len(pRs!NroDocIDContr) < 15, pRs!NroDocIDContr & String(Abs(15 - Len(pRs!NroDocIDContr)), " "), _
'                   pRs!NroDocIDContr))
'             sCad1 = sCad1 & String(20, " ")
'             sCad1 = sCad1 & String(20, " ")
'             sCad1 = sCad1 & String(30, " ")
'             sCad1 = sCad1 & String(8, " ")
'             sCad1 = sCad1 & pRs!SexoContr
'             sCad1 = sCad1 & IIf(Len(pRs!DirContr) > 50, Mid(pRs!DirContr, 1, 50), _
'                   IIf(Len(pRs!DirContr) < 50, pRs!DirContr & String(Abs(50 - Len(pRs!DirContr)), " "), _
'                   pRs!DirContr))
'             sCad1 = sCad1 & Format(pRs!UbiDepContr, "00")
'             sCad1 = sCad1 & Format(pRs!UbiProvContr, "00")
'             sCad1 = sCad1 & Format(pRs!UbiDistrContr, "00")
'             sCad1 = sCad1 & String(12, " ")
'             sCad1 = sCad1 & String(12, " ")
'             sCad1 = sCad1 & String(1, " ")
'             sCad1 = sCad1 & IIf(Len(pRs!RazonSocContr) > 50, Mid(pRs!RazonSocContr, 1, 50), _
'                   IIf(Len(pRs!RazonSocContr) < 50, pRs!RazonSocContr & String(Abs(50 - Len(pRs!RazonSocContr)), " "), _
'                   pRs!RazonSocContr))
'             sCad1 = sCad1 & String(30, " ")
'             sCad1 = sCad1 & String(20, " ")
'             sTelefono = pRs!TelfContr
'             sTelefono = QuitarCaracter(QuitarCaracter(QuitarCaracter(sTelefono, "-"), ")"), "(")
'             sCad1 = sCad1 & IIf(Len(sTelefono) > 12, Mid(sTelefono, 1, 12), _
'                   IIf(Len(sTelefono) < 12, sTelefono & String(Abs(12 - Len(sTelefono)), " "), _
'                   sTelefono))
'             sCad1 = sCad1 & Format(pRs!NroCertContr, "0000000000")
'             Print #1, ""; sCad1
             
              'ASEGURADO
             sCad2 = "2" & pRs!TpoPersAseg
             sCad2 = sCad2 & Format(pRs!TpoDocIDAseg, "0")
             sCad2 = sCad2 & Format(pRs!NroDocIDAseg, "000000000000000")
             sCad2 = sCad2 & IIf(Len(pRs!ApePatAseg) > 20, Mid(pRs!ApePatAseg, 1, 20), _
                   IIf(Len(pRs!ApePatAseg) < 20, pRs!ApePatAseg & String(Abs(20 - Len(pRs!ApePatAseg)), " "), _
                   pRs!ApePatAseg))
             sCad2 = sCad2 & IIf(Len(pRs!ApeMatAseg) > 20, Mid(pRs!ApeMatAseg, 1, 20), _
                   IIf(Len(pRs!ApeMatAseg) < 20, pRs!ApeMatAseg & String(Abs(20 - Len(pRs!ApeMatAseg)), " "), _
                   pRs!ApeMatAseg))
             sCad2 = sCad2 & IIf(Len(pRs!NombreAseg) > 30, Mid(pRs!NombreAseg, 1, 30), _
                   IIf(Len(pRs!NombreAseg) < 30, pRs!NombreAseg & String(Abs(30 - Len(pRs!NombreAseg)), " "), _
                   pRs!NombreAseg))
             sCad2 = sCad2 & Format(pRs!FecNacAseg, "yyyymmdd")
             sCad2 = sCad2 & pRs!SexoAseg
             sCad2 = sCad2 & IIf(Len(pRs!DirAseg) > 50, Mid(pRs!DirAseg, 1, 50), _
                   IIf(Len(pRs!DirAseg) < 50, pRs!DirAseg & String(Abs(50 - Len(pRs!DirAseg)), " "), _
                   pRs!DirAseg))
             sCad2 = sCad2 & Format(pRs!UbiDepAseg, "00")
             sCad2 = sCad2 & Format(pRs!UbiProvAseg, "00")
             sCad2 = sCad2 & Format(pRs!UbiDistrAseg, "00")
             'Quitar caracteres al telefono como "(",")" y/o "-"
             sTelefono = pRs!TelfAseg
             sTelefono = QuitarCaracter(QuitarCaracter(QuitarCaracter(sTelefono, "-"), ")"), "(")
             sCad2 = sCad2 & IIf(Len(sTelefono) > 12, Mid(sTelefono, 1, 12), _
                   IIf(Len(sTelefono) < 12, String(Abs(12 - Len(sTelefono)), "0") & sTelefono, _
                   sTelefono))
             sTelefono = pRs!CelAseg
             sTelefono = QuitarCaracter(QuitarCaracter(QuitarCaracter(sTelefono, "-"), ")"), "(")
             sCad2 = sCad2 & IIf(Len(sTelefono) > 12, Mid(sTelefono, 1, 12), _
                   IIf(Len(sTelefono) < 12, String(Abs(12 - Len(sTelefono)), "0") & sTelefono, _
                   sTelefono))
             sCad2 = sCad2 & pRs!EstCivAseg
             sCad2 = sCad2 & String(1, "0")
             sCad2 = sCad2 & Format(QuitarCaracter(pRs!NroCertAseg, "-"), "0000000000")
             Print #1, ""; sCad2
            End If
            
            If (pRs!NroDocIDBenf) <> "" Then
                'BENEFICIARIO
                sCad3 = "3" & Format(pRs!NroDocIDBenf, "000000000000000")
                sCad3 = sCad3 & IIf(Len(pRs!ApePatBenf) > 20, Mid(pRs!ApePatBenf, 1, 20), _
                      IIf(Len(pRs!ApePatBenf) < 20, pRs!ApePatBenf & String(Abs(20 - Len(pRs!ApePatBenf)), " "), _
                      pRs!ApePatBenf))
                sCad3 = sCad3 & IIf(Len(pRs!ApeMatBenf) > 20, Mid(pRs!ApeMatBenf, 1, 20), _
                      IIf(Len(pRs!ApeMatBenf) < 20, pRs!ApeMatBenf & String(Abs(20 - Len(pRs!ApeMatBenf)), " "), _
                      pRs!ApeMatBenf))
                sCad3 = sCad3 & IIf(Len(pRs!NombreBenf) > 30, Mid(pRs!NombreBenf, 1, 30), _
                      IIf(Len(pRs!NombreBenf) < 30, pRs!NombreBenf & String(Abs(30 - Len(pRs!NombreBenf)), " "), _
                      pRs!NombreBenf))
                sCad3 = sCad3 & Format(pRs!FecNacBenf, "yyyymmdd")
                sCad3 = sCad3 & IIf(Len(pRs!ParentBenf) > 2, Mid(pRs!ParentBenf, 1, 2), _
                      IIf(Len(pRs!ParentBenf) < 2, pRs!ParentBenf & String(Abs(2 - Len(pRs!ParentBenf)), " "), _
                      pRs!ParentBenf))
                sCad3 = sCad3 & Format(pRs!PorcBenf, "0000000000000")
                sCad3 = sCad3 & Format(QuitarCaracter(pRs!NroCertBenf, "-"), "0000000000")
                Print #1, ""; sCad3
            End If
            If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
                If (pRs!NroDocIDCony) <> "" Then
                      'CONYUGE
                      sCad4 = "4" & Format(pRs!NroDocIDCony, "000000000000000")
                      sCad4 = sCad4 & IIf(Len(pRs!ApePatCony) > 20, Mid(pRs!ApePatCony, 1, 20), _
                            IIf(Len(pRs!ApePatCony) < 20, pRs!ApePatCony & String(Abs(20 - Len(pRs!ApePatCony)), " "), _
                            pRs!ApePatCony))
                      sCad4 = sCad4 & IIf(Len(pRs!ApeMatCony) > 20, Mid(pRs!ApeMatCony, 1, 20), _
                            IIf(Len(pRs!ApeMatCony) < 20, pRs!ApeMatCony & String(Abs(20 - Len(pRs!ApeMatCony)), " "), _
                            pRs!ApeMatCony))
                      sCad4 = sCad4 & IIf(Len(pRs!NombreCony) > 30, Mid(pRs!NombreCony, 1, 30), _
                            IIf(Len(pRs!NombreCony) < 30, pRs!NombreCony & String(Abs(30 - Len(pRs!NombreCony)), " "), _
                            pRs!NombreCony))
                      sCad4 = sCad4 & Format(pRs!FecNacCony, "yyyymmdd")
                      sCad4 = sCad4 & Format(QuitarCaracter(pRs!NroCertCony, "-"), "0000000000")
                      Print #1, ""; sCad4
                End If
            End If
          sCtaCod = pRs!cCtaCod
          pRs.MoveNext
        Next nConteo
    End If
ElseIf pnTipo = "2" Then
If Not (pRs.EOF And pRs.BOF) Then
        sCtaCod = ""
        
        For nConteo = 1 To pRs.RecordCount
          If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
            nContMat = 1
             'DECLARACION
             sCad0 = "0" & Format(QuitarCaracter(pRs!NumCert, "-"), "000000000000")
             sCad0 = sCad0 & Format(pRs!FecVigInicial, "yyyymmdd")
             sCad0 = sCad0 & Format(pRs!FecVigFinal, "yyyymmdd")
             sCad0 = sCad0 & Format(pRs!PrimaMensual * 100, "0000000000000")
             sCad0 = sCad0 & Format(pRs!CodPlan, "00000")
             Print #1, ""; sCad0
             
            'DECLRACION JURADA
             sCad1 = "1" & pRs!TpoDJ
             sCad1 = sCad1 & Format(QuitarCaracter(pRs!NroDJ, "-"), "000000000000000")
             sCad1 = sCad1 & pRs!TpoGarDJ
             sCad1 = sCad1 & Format(pRs!FechaDJ, "yyyymmdd")
             sCad1 = sCad1 & IIf(Len(pRs!AntNroDJ) > 15, Mid(pRs!AntNroDJ, 1, 15), _
                   IIf(Len(pRs!AntNroDJ) < 15, pRs!AntNroDJ & String(Abs(15 - Len(pRs!AntNroDJ)), " "), _
                   pRs!AntNroDJ))
             sCad1 = sCad1 & IIf(Len(pRs!NomCliDJ) > 50, Mid(pRs!NomCliDJ, 1, 50), _
                   IIf(Len(pRs!NomCliDJ) < 50, pRs!NomCliDJ & String(Abs(50 - Len(pRs!NomCliDJ)), " "), _
                   pRs!NomCliDJ))
             sCad1 = sCad1 & Format(IIf(Trim(pRs!DniCliDJ) = "", "0", pRs!DniCliDJ), "000000000000")
             sCad1 = sCad1 & IIf(Len(pRs!NomConyDJ) > 50, Mid(pRs!NomConyDJ, 1, 50), _
                   IIf(Len(pRs!NomConyDJ) < 50, pRs!NomConyDJ & String(Abs(50 - Len(pRs!NomConyDJ)), " "), _
                   pRs!NomConyDJ))
             sCad1 = sCad1 & Format(IIf(Trim(pRs!DniConyDJ) = "", "0", pRs!DniConyDJ), "000000000000")
             sCad1 = sCad1 & IIf(Len(pRs!DirDJ) > 50, Mid(pRs!DirDJ, 1, 50), _
                   IIf(Len(pRs!DirDJ) < 50, pRs!DirDJ & String(Abs(50 - Len(pRs!DirDJ)), " "), _
                   pRs!DirDJ))
             sCad1 = sCad1 & pRs!MonedaDJ
             sCad1 = sCad1 & Format(pRs!MontoDJ * 100, "0000000000000")
             sCad1 = sCad1 & Format(QuitarCaracter(pRs!NroCertDJ, "-"), "000000000000")
             Print #1, ""; sCad1
             
             'MATERIAS ASEGURADAS
             sCad2 = "2" & Format(QuitarCaracter(pRs!NroDJMatAseg, "-"), "000000000000000")
             sCad2 = sCad2 & pRs!TpoGarMatAseg
             sCad2 = sCad2 & Format(nContMat, "00") 'Format(pRs!ContMatAseg, "00")
             sCad2 = sCad2 & IIf(Len(pRs!DirMatAseg) > 50, Mid(pRs!DirMatAseg, 1, 50), _
                   IIf(Len(pRs!DirMatAseg) < 50, pRs!DirMatAseg & String(Abs(50 - Len(pRs!DirMatAseg)), " "), _
                   pRs!DirMatAseg))
             sCad2 = sCad2 & Format(pRs!UbiDepMatAseg, "00")
             sCad2 = sCad2 & Format(pRs!UbiProvMatAseg, "00")
             sCad2 = sCad2 & Format(pRs!UbiDistMatAseg, "00")
             sCad2 = sCad2 & Format(pRs!AnoConsMatAseg, "0000")
             sCad2 = sCad2 & IIf(Len(pRs!NroPisMatAseg) > 2, Mid(pRs!NroPisMatAseg, 1, 2), _
                        IIf(Len(pRs!NroPisMatAseg) < 2, pRs!NroPisMatAseg & String(Abs(2 - Len(pRs!NroPisMatAseg)), " "), _
                        pRs!NroPisMatAseg))
             sCad2 = sCad2 & pRs!UsoMatAseg
             sCad2 = sCad2 & IIf(pRs!MatMatAseg = "", " ", pRs!MatMatAseg)
             sCad2 = sCad2 & pRs!TpoMatAseg
             sCad2 = sCad2 & IIf(Len(pRs!DescMatAseg) > 100, Mid(pRs!DescMatAseg, 1, 100), _
                   IIf(Len(pRs!DescMatAseg) < 100, pRs!DescMatAseg & String(Abs(100 - Len(pRs!DescMatAseg)), " "), _
                   pRs!DescMatAseg))
             sCad2 = sCad2 & Format(QuitarCaracter(pRs!NroCertMatAseg, "-"), "000000000000")
             Print #1, ""; sCad2
             
              'ASEGURADO
             sCad3 = "4" & Format(pRs!TpoDocIDAseg, "0")
             sCad3 = sCad3 & Format(QuitarCaracter(pRs!NroDocIDAseg, "-"), "000000000000000")
             sCad3 = sCad3 & IIf(Len(pRs!ApePatAseg) > 20, Mid(pRs!ApePatAseg, 1, 20), _
                   IIf(Len(pRs!ApePatAseg) < 20, pRs!ApePatAseg & String(Abs(20 - Len(pRs!ApePatAseg)), " "), _
                   pRs!ApePatAseg))
             sCad3 = sCad3 & IIf(Len(pRs!ApeMatAseg) > 20, Mid(pRs!ApeMatAseg, 1, 20), _
                   IIf(Len(pRs!ApeMatAseg) < 20, pRs!ApeMatAseg & String(Abs(20 - Len(pRs!ApeMatAseg)), " "), _
                   pRs!ApeMatAseg))
             sCad3 = sCad3 & IIf(Len(pRs!NombreAseg) > 30, Mid(pRs!NombreAseg, 1, 30), _
                   IIf(Len(pRs!NombreAseg) < 30, pRs!NombreAseg & String(Abs(30 - Len(pRs!NombreAseg)), " "), _
                   pRs!NombreAseg))
             sCad3 = sCad3 & Format(pRs!FecNacAseg, "yyyymmdd")
             sCad3 = sCad3 & pRs!SexoAseg
             sCad3 = sCad3 & IIf(Len(pRs!DirAseg) > 50, Mid(pRs!DirAseg, 1, 50), _
                   IIf(Len(pRs!DirAseg) < 50, pRs!DirAseg & String(Abs(50 - Len(pRs!DirAseg)), " "), _
                   pRs!DirAseg))
             sCad3 = sCad3 & Format(pRs!UbiDepAseg, "00")
             sCad3 = sCad3 & Format(pRs!UbiProvAseg, "00")
             sCad3 = sCad3 & Format(pRs!UbiDistrAseg, "00")
             sTelefono = pRs!TelfAseg
             sTelefono = QuitarCaracter(QuitarCaracter(QuitarCaracter(sTelefono, "-"), ")"), "(")
             sCad3 = sCad3 & IIf(Len(sTelefono) > 12, Mid(sTelefono, 1, 12), _
                   IIf(Len(sTelefono) < 12, sTelefono & String(Abs(12 - Len(sTelefono)), " "), _
                   sTelefono))
             sCad3 = sCad3 & pRs!EstCivAseg
             sCad3 = sCad3 & Format(QuitarCaracter(pRs!NroCertAseg, "-"), "000000000000")
             Print #1, ""; sCad3
          End If
          
          If Trim(sCtaCod) = Trim(pRs!cCtaCod) Then
            'MATERIAS ASEGURADAS
             sCad4 = "2" & Format(QuitarCaracter(pRs!NroDJMatAseg, "-"), "000000000000000")
             sCad4 = sCad4 & pRs!TpoGarMatAseg
             sCad4 = sCad4 & Format(nContMat, "00") 'Format(pRs!ContMatAseg, "00")
             sCad4 = sCad4 & IIf(Len(pRs!DirMatAseg) > 50, Mid(pRs!DirMatAseg, 1, 50), _
                   IIf(Len(pRs!DirMatAseg) < 50, pRs!DirMatAseg & String(Abs(50 - Len(pRs!DirMatAseg)), " "), _
                   pRs!DirMatAseg))
             sCad4 = sCad4 & Format(pRs!UbiDepMatAseg, "00")
             sCad4 = sCad4 & Format(pRs!UbiProvMatAseg, "00")
             sCad4 = sCad4 & Format(pRs!UbiDistMatAseg, "00")
             sCad4 = sCad4 & Format(pRs!AnoConsMatAseg, "0000")
             sCad4 = sCad4 & IIf(Len(pRs!NroPisMatAseg) > 2, Mid(pRs!NroPisMatAseg, 1, 2), _
                        IIf(Len(pRs!NroPisMatAseg) < 2, pRs!NroPisMatAseg & String(Abs(2 - Len(pRs!NroPisMatAseg)), " "), _
                        pRs!NroPisMatAseg))
             sCad4 = sCad4 & pRs!UsoMatAseg
             sCad4 = sCad4 & IIf(pRs!MatMatAseg = "", " ", pRs!MatMatAseg)
             sCad4 = sCad4 & pRs!TpoMatAseg
             sCad4 = sCad4 & IIf(Len(pRs!DescMatAseg) > 100, Mid(pRs!DescMatAseg, 1, 100), _
                   IIf(Len(pRs!DescMatAseg) < 100, pRs!DescMatAseg & String(Abs(100 - Len(pRs!DescMatAseg)), " "), _
                   pRs!DescMatAseg))
             sCad4 = sCad4 & Format(QuitarCaracter(pRs!NroCertMatAseg, "-"), "000000000000")
            Print #1, ""; sCad4
          End If
          sCtaCod = pRs!cCtaCod
          nContMat = nContMat + 1
          pRs.MoveNext
        Next nConteo
    End If
End If

End Sub
Private Function QuitarCaracter(ByVal psCadena As String, ByVal psCaracter As String) As String
Dim nPosicion As Integer
Dim nTamano As Integer
Dim sResultado As String
sResultado = psCadena
nTamano = Len(psCadena)
nPosicion = InStr(psCadena, psCaracter)
Do While nPosicion <> 0
    sResultado = Mid(sResultado, 1, nPosicion - 1) & Mid(sResultado, nPosicion + 1, nTamano - nPosicion)
    nPosicion = InStr(sResultado, psCaracter)
Loop
QuitarCaracter = sResultado
End Function
Private Function ObtenerMonto(pRs As ADODB.Recordset) As Double
Dim i As Integer
Dim nMontoFinal As Double
Dim sCtaCod As String
sCtaCod = ""
For i = 0 To pRs.RecordCount - 1
    If sTipoD = "AF" Then
        If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
           nMontoFinal = nMontoFinal + (CDbl(pRs!PrimaMensual))
        End If
        sCtaCod = pRs!cCtaCod
    ElseIf sTipoD = "CA" Then
        nMontoFinal = nMontoFinal + CDbl(pRs!PrimaMensual)
    End If
    pRs.MoveNext
Next i
ObtenerMonto = nMontoFinal
End Function

Private Function ObtenerNumMov(pRs As ADODB.Recordset, ByVal pnVerDato As Integer) As Integer
Dim i As Integer
Dim nNumMov As Integer
Dim nNumBenf As Integer
Dim nNumCony As Integer
Dim nNumRep As Integer
Dim sCtaCod As String
pRs.MoveFirst
sCtaCod = ""
For i = 0 To pRs.RecordCount - 1
    If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
        nNumMov = nNumMov + 1
    Else
        nNumRep = nNumRep + 1
    End If
    If sTipoTrama = "1" Then
        If (pRs!NroDocIDBenf) <> "" Then
            nNumBenf = nNumBenf + 1
        End If
        If Trim(sCtaCod) <> Trim(pRs!cCtaCod) Then
            If (pRs!NroDocIDCony) <> "" Then
                nNumCony = nNumCony + 1
            End If
        End If
    End If
    sCtaCod = pRs!cCtaCod
    pRs.MoveNext
Next i

If pnVerDato = 1 Then
    ObtenerNumMov = nNumMov
ElseIf pnVerDato = 2 Then
    ObtenerNumMov = nNumCony
ElseIf pnVerDato = 3 Then
    ObtenerNumMov = nNumBenf
ElseIf pnVerDato = 4 Then
    ObtenerNumMov = nNumRep
End If
End Function


Private Sub EstructuraCargos(ByVal pRs As ADODB.Recordset, ByVal pnTipo As Integer)
Dim sCad0 As String
Dim sCad1 As String
Dim sCad2 As String
Dim sCad3 As String
Dim nConteo As Integer
Dim sTelefono As String
pRs.MoveFirst
If pnTipo = "1" Then
    'DEPOSITO
    sCad0 = "D" & Format(nMonto * 100, "0000000000000")
    sCad0 = sCad0 & sTipoPago
    sCad0 = sCad0 & IIf(Len(txtCta.Text) > 20, Mid(txtCta.Text, 1, 20), _
                IIf(Len(txtCta.Text) < 20, txtCta.Text & String(Abs(20 - Len(txtCta.Text)), " "), _
                txtCta.Text))
    sCad0 = sCad0 & Format(Me.txtMesDepostito.Text, "yyyymmdd")
    sCad0 = sCad0 & IIf(Len(txtBanco.Text) > 60, Mid(txtBanco.Text, 1, 60), _
            IIf(Len(txtBanco.Text) < 60, txtBanco.Text & String(Abs(60 - Len(txtBanco.Text)), " "), _
            txtBanco.Text))
    sCad0 = sCad0 & IIf(Len(txtNroDocumento.Text) > 16, Mid(txtNroDocumento.Text, 1, 16), _
            IIf(Len(txtNroDocumento.Text) < 16, txtNroDocumento.Text & String(Abs(16 - Len(txtNroDocumento.Text)), " "), _
            txtNroDocumento.Text))
    sCad0 = sCad0 & IIf(Len(txtObservaciones.Text) > 30, Mid(txtObservaciones.Text, 1, 30), _
            IIf(Len(txtObservaciones.Text) < 30, txtObservaciones.Text & String(Abs(30 - Len(txtObservaciones.Text)), " "), _
            txtObservaciones.Text))
    Print #1, ""; sCad0
    
    If Not (pRs.EOF And pRs.BOF) Then
        For nConteo = 1 To pRs.RecordCount
                    
         'RECAUDACION
          sCad1 = "C" & IIf(Len(pRs!CtaCli) > 20, Mid(pRs!CtaCli, 1, 20), _
                IIf(Len(pRs!CtaCli) < 20, pRs!CtaCli & String(Abs(20 - Len(pRs!CtaCli)), " "), _
                pRs!CtaCli))
          sCad1 = sCad1 & IIf(Len(pRs!NombreCli) > 80, Mid(pRs!NombreCli, 1, 80), _
                IIf(Len(pRs!NombreCli) < 80, pRs!NombreCli & String(Abs(80 - Len(pRs!NombreCli)), " "), _
                pRs!NombreCli))
          sCad1 = sCad1 & Format(CDbl(pRs!PrimaMensual) * 100, "0000000000000")
          sCad1 = sCad1 & Format(CDbl(pRs!PrimaMensualB) * 100, "0000000000000")
          sCad1 = sCad1 & Format(QuitarCaracter(pRs!NroCert, "-"), "0000000000")
          sCad1 = sCad1 & IIf(Len(pRs!Observaciones) > 30, Mid(pRs!Observaciones, 1, 30), _
                IIf(Len(pRs!Observaciones) < 30, pRs!Observaciones & String(Abs(30 - Len(pRs!Observaciones)), " "), _
                pRs!Observaciones))
          Print #1, ""; sCad1
          
                  
          pRs.MoveNext
        Next nConteo
    End If
ElseIf pnTipo = "2" Then
    'DEPOSITO
    sCad0 = "D" & Format(nMonto * 100, "0000000000000")
    sCad0 = sCad0 & sTipoPago
    sCad0 = sCad0 & IIf(Len(txtCta.Text) > 20, Mid(txtCta.Text, 1, 20), _
                IIf(Len(txtCta.Text) < 20, txtCta.Text & String(Abs(20 - Len(txtCta.Text)), " "), _
                txtCta.Text))
    sCad0 = sCad0 & Format(Me.txtMesDepostito.Text, "yyyymmdd")
    sCad0 = sCad0 & IIf(Len(txtBanco.Text) > 60, Mid(txtBanco.Text, 1, 60), _
            IIf(Len(txtBanco.Text) < 60, txtBanco.Text & String(Abs(60 - Len(txtBanco.Text)), " "), _
            txtBanco.Text))
    sCad0 = sCad0 & IIf(Len(txtNroDocumento.Text) > 16, Mid(txtNroDocumento.Text, 1, 16), _
            IIf(Len(txtNroDocumento.Text) < 16, txtNroDocumento.Text & String(Abs(16 - Len(txtNroDocumento.Text)), " "), _
            txtNroDocumento.Text))
    sCad0 = sCad0 & IIf(Len(txtObservaciones.Text) > 30, Mid(txtObservaciones.Text, 1, 30), _
            IIf(Len(txtObservaciones.Text) < 30, txtObservaciones.Text & String(Abs(30 - Len(txtObservaciones.Text)), " "), _
            txtObservaciones.Text))
    Print #1, ""; sCad0
    
    If Not (pRs.EOF And pRs.BOF) Then
        For nConteo = 1 To pRs.RecordCount
         'RECAUDACION
          sCad1 = "C" & IIf(Len(pRs!CtaCli) > 20, Mid(pRs!CtaCli, 1, 20), _
                IIf(Len(pRs!CtaCli) < 20, pRs!CtaCli & String(Abs(20 - Len(pRs!CtaCli)), " "), _
                pRs!CtaCli))
          sCad1 = sCad1 & IIf(Len(pRs!NombreCli) > 80, Mid(pRs!NombreCli, 1, 80), _
                IIf(Len(pRs!NombreCli) < 80, pRs!NombreCli & String(Abs(80 - Len(pRs!NombreCli)), " "), _
                pRs!NombreCli))
          sCad1 = sCad1 & Format(CDbl(pRs!PrimaMensual) * 100, "0000000000000")
          sCad1 = sCad1 & Format(CDbl(pRs!PrimaMensualB) * 100, "0000000000000")
          sCad1 = sCad1 & Format(QuitarCaracter(pRs!NroCert, "-"), "000000000000")
          sCad1 = sCad1 & IIf(Len(pRs!Observaciones) > 30, Mid(pRs!Observaciones, 1, 30), _
                IIf(Len(pRs!Observaciones) < 30, pRs!Observaciones & String(Abs(30 - Len(pRs!Observaciones)), " "), _
                pRs!Observaciones))
          Print #1, ""; sCad1
          pRs.MoveNext
        Next nConteo
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sMonedaTXT = ""
sTipoD = ""
sTipoTrama = ""
sTipoPago = ""
sMes = ""
sAno = ""
bValidaFec = False
dfechaini = ""
dfechafin = ""
psArchivoAGrabarC = ""
nMonto = 0
End Sub

Private Sub txtMesActual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call ValidaFecha
End If
End Sub

Private Function validaDatos() As Boolean

    If Me.txtCodPatrocinador.Text = "" Then
        MsgBox "Ingrese Codigo Patrocinador.", vbInformation, "Aviso"
        validaDatos = False
        txtCodPatrocinador.SetFocus
        Exit Function
    End If
    
    If Me.txtNumRamo.Text = "" Then
        MsgBox "Ingrese Nro. Ramo.", vbInformation, "Aviso"
        validaDatos = False
        txtNumRamo.SetFocus
        Exit Function
    End If
    
    If Me.txtCodProducto.Text = "" Then
        MsgBox "Ingrese Codigo Producto.", vbInformation, "Aviso"
        validaDatos = False
        txtCodProducto.SetFocus
        Exit Function
    End If
    
    If Me.txtNumPoliza.Text = "" Then
        MsgBox "Ingrese Nro. Poliza.", vbInformation, "Aviso"
        validaDatos = False
        txtNumPoliza.SetFocus
        Exit Function
    End If
    
    Call ValidaFecha
    If Not bValidaFec Then
        Me.txtMesActual.SetFocus
        validaDatos = False
        Exit Function
    End If
    
    If Me.cmdTipo.Text = "" Then
        MsgBox "Seleccione Tipo de Seguro.", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
    If Me.cmbTipoDeclaracion.Text = "" Then
        MsgBox "Seleccione Tipo de Declaracion.", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
    If Me.cmbMoneda.Text = "" Then
        MsgBox "Seleccione la moneda", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
    If Me.cmbAgencia.Text = "" Then
        MsgBox "Seleccione la Agencia.", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
     If Me.txtCodPlan.Text = "" Then
        MsgBox "Ingrese Codigo del Plan.", vbInformation, "Aviso"
        validaDatos = False
        txtCodPatrocinador.SetFocus
        Exit Function
    End If
    
    If sTipoTrama = "1" Then
        If Trim(Right(Me.cmbMicroseguros.Text, 2)) = "0" Then
        MsgBox "Seleccione Tipo de Microseguro", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
    End If
    If sTipoD = "CA" Then
        If Me.cmbTipoPago.Text = "" Then
        MsgBox "Seleccione Tipo Pago.", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
        End If
        
        If Me.txtCta.Text = "" Then
        MsgBox "Ingrese Cuenta.", vbInformation, "Aviso"
        validaDatos = False
        txtCta.SetFocus
        Exit Function
        End If
        
        If Me.txtBanco.Text = "" Then
        MsgBox "Ingrese Nombre del Banco.", vbInformation, "Aviso"
        validaDatos = False
        txtBanco.SetFocus
        Exit Function
        End If
        
        If Me.txtNroDocumento.Text = "" Then
        MsgBox "Ingrese Nro. Documento del Deposito.", vbInformation, "Aviso"
        validaDatos = False
        txtNroDocumento.SetFocus
        Exit Function
        End If
        
        If Me.txtMesDepostito.Text = "__/__/____" Then
        MsgBox "Ingrese fecha de deposito.", vbInformation, "Aviso"
        validaDatos = False
        txtMesDepostito.SetFocus
        Exit Function
        End If
        
        Dim nPosicion As Integer
        nPosicion = InStr(Me.txtMesDepostito.Text, "_")
        If nPosicion > 0 Then
            MsgBox "Error en fecha de deposito.", vbInformation, "Aviso"
            validaDatos = False
            txtMesDepostito.SetFocus
            Exit Function
        End If
        If CInt(Mid(Me.txtMesDepostito.Text, 7, 4)) > CInt(Mid(gdFecSis, 7, 4)) Or CInt(Mid(Me.txtMesDepostito.Text, 7, 4)) < (CInt(Mid(gdFecSis, 7, 4)) - 100) Then
            MsgBox "Año fuera del rango.", vbInformation, "Aviso"
            validaDatos = False
            txtMesDepostito.SetFocus
            Exit Function
        End If
        
        If CInt(Mid(Me.txtMesDepostito.Text, 4, 2)) > 12 Or CInt(Mid(Me.txtMesDepostito.Text, 4, 2)) < 1 Then
            MsgBox "Error en Mes.", vbInformation, "Aviso"
            validaDatos = False
            txtMesDepostito.SetFocus
            Exit Function
        End If
            
        Dim nDiasEnMes As Integer
        nDiasEnMes = CInt(DateDiff("d", CDate("01" & Mid(Me.txtMesDepostito.Text, 3, 8)), DateAdd("M", 1, CDate("01" & Mid(Me.txtMesDepostito.Text, 3, 8)))))
        
        If CInt(Mid(Me.txtMesDepostito.Text, 1, 2)) > nDiasEnMes Or CInt(Mid(Me.txtMesDepostito.Text, 1, 2)) < 1 Then
            MsgBox "Dia fuera del rango.", vbInformation, "Aviso"
            validaDatos = False
            txtMesDepostito.SetFocus
            Exit Function
        End If
    End If
validaDatos = True
End Function

Private Sub ValidaFecha()
If Me.txtMesActual.Text <> "__/____" Then
        sMes = Mid(Me.txtMesActual.Text, 1, 2)
        sAno = Mid(Me.txtMesActual.Text, 4, 4)
        Dim nPosAno As Integer
        Dim nPosMes As Integer
        If sAno = "____" Then
            bValidaFec = False
            Me.lblMesActual.Caption = ""
            MsgBox "Ingrese Año.", vbInformation, "Aviso"
        Else
            nPosAno = InStr(1, Trim(sAno), "_")
            If nPosAno > 0 Then
                bValidaFec = False
                Me.lblMesActual.Caption = ""
                MsgBox "Error de Año.", vbInformation, "Aviso"
            Else
                If (CDbl(Mid(CStr(gdFecSis), 7, 4)) >= CDbl(sAno)) And ((CDbl(Mid(CStr(gdFecSis), 7, 4)) - 100) <= CDbl(sAno)) Then
                    nPosMes = InStr(1, Trim(sMes), "_")
                    If nPosMes > 0 Then
                        bValidaFec = False
                        Me.lblMesActual.Caption = ""
                        MsgBox "Error de Mes.", vbInformation, "Aviso"
                    Else
                        If (CDbl(Mid(CStr(gdFecSis), 7, 4)) = CDbl(sAno)) And (CDbl(Mid(CStr(gdFecSis), 4, 2)) < CDbl(sMes)) Then
                            bValidaFec = False
                            Me.lblMesActual.Caption = ""
                            MsgBox "Mes no puede ser mayor al Actual.", vbInformation, "Aviso"
                        ElseIf (CDbl(Mid(CStr(gdFecSis), 7, 4)) = CDbl(sAno)) And (CDbl(Mid(CStr(gdFecSis), 4, 2)) = CDbl(sMes)) Then
                            bValidaFec = False
                            Me.lblMesActual.Caption = ""
                            MsgBox "Mes no puede ser igual al actual.", vbInformation, "Aviso"
                        Else
                            If sMes = "01" Then
                                Me.lblMesActual.Caption = "Enero " & sAno
                                bValidaFec = True
                            ElseIf sMes = "02" Then
                                Me.lblMesActual.Caption = "Febrero " & sAno
                                bValidaFec = True
                            ElseIf sMes = "03" Then
                                Me.lblMesActual.Caption = "Marzo " & sAno
                                bValidaFec = True
                            ElseIf sMes = "04" Then
                                Me.lblMesActual.Caption = "Abril " & sAno
                                bValidaFec = True
                            ElseIf sMes = "05" Then
                                Me.lblMesActual.Caption = "Mayo " & sAno
                                bValidaFec = True
                            ElseIf sMes = "06" Then
                                Me.lblMesActual.Caption = "Junio " & sAno
                                bValidaFec = True
                            ElseIf sMes = "07" Then
                                Me.lblMesActual.Caption = "Julio " & sAno
                                bValidaFec = True
                            ElseIf sMes = "08" Then
                                Me.lblMesActual.Caption = "Agosto " & sAno
                                bValidaFec = True
                            ElseIf sMes = "09" Then
                                Me.lblMesActual.Caption = "Septiembre " & sAno
                                bValidaFec = True
                            ElseIf sMes = "10" Then
                                Me.lblMesActual.Caption = "Octubre " & sAno
                                bValidaFec = True
                            ElseIf sMes = "11" Then
                                Me.lblMesActual.Caption = "Noviembre " & sAno
                                bValidaFec = True
                            ElseIf sMes = "12" Then
                                Me.lblMesActual.Caption = "Diciembre " & sAno
                                bValidaFec = True
                            Else
                                bValidaFec = False
                                MsgBox "Error de Mes.", vbInformation, "Aviso"
                            End If
                        End If
                    End If
                Else
                    If CDbl(Mid(CStr(gdFecSis), 7, 4)) < CDbl(sAno) Then
                        bValidaFec = False
                        Me.lblMesActual.Caption = ""
                        MsgBox "Año no puede ser mayor al actual.", vbInformation, "Aviso"
                    Else
                        bValidaFec = False
                        Me.lblMesActual.Caption = ""
                        MsgBox "Año fuera de rango.", vbInformation, "Aviso"
                    End If
                End If
            End If
        End If
    Else
        bValidaFec = False
        Me.lblMesActual.Caption = ""
        MsgBox "Ingrese Mes y Año.", vbInformation, "Aviso"
    End If
End Sub
