VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPITReporteOpeInterCMAC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Operaciones InterCajas"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo de Reporte "
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4215
      Begin VB.OptionButton optTipo 
         Caption         =   "Resumen de operaciones"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Detallado"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   " Realizados entre las fechas "
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4215
      Begin MSMask.MaskEdBox mskFechaDe 
         Height          =   300
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaHasta 
         Height          =   300
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "De: "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta: "
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   3240
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   3240
      Width           =   1230
   End
   Begin VB.Frame fraModalidad 
      Caption         =   " Modo de Operaciónes "
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton optModo 
         Caption         =   "De CMAC Maynas"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optModo 
         Caption         =   "De otras CMAC"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPITReporteOpeInterCMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim lsTitulo As String
Dim lsSubTitulo As String, lsSubTitulo2 As String
Dim lsReporte As String
Dim lrsReporte As ADODB.Recordset

    lsTitulo = "Reporte de Operaciones InterCMAC"
    If optTipo(0).Value Then
        lsTitulo = lsTitulo & " - Detallado"
    Else
        lsTitulo = lsTitulo & " - Resumen"
    End If
    
    
    If optModo(0).Value Then
        lsSubTitulo = " Clientes de CMAC Maynas "
    Else
        lsSubTitulo = " Clientes de Otras CMACS "
    End If
    
    lsSubTitulo2 = "Realizadas entre las fechas de " & Format(mskFechaDe.Text, "DD/MM/YYYY") & " al " & Format(mskFechaHasta.Text, "DD/MM/YYYY")
    
    lsSubTitulo = lsSubTitulo & " - " & lsSubTitulo2
    
    If Not IsDate(Me.mskFechaDe.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFechaDe.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaHasta.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFechaHasta.SetFocus
        Exit Sub
    End If
    
    
    lsReporte = lsReporte & Space(5) & fijarTamanoTexto("CMAC MAYNAS S.A.", 26) & Space(40) & "FECHA   : " & Format(Now(), "DD/MM/YYYY HH:MM:SS") & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto(gsNomAge, 26) & Space(40) & "USUARIO : " & gsCodUser & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto(lsTitulo, 100, 2) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto(lsSubTitulo, 100, 2) & Chr(10)
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    
    If optModo(0).Value Then 'Recepcion
        If optTipo(0).Value Then 'Detallade
            lsReporte = lsReporte & reporteOpeInterCMAC_Recepcion_Detalle(mskFechaDe.Text, mskFechaHasta.Text)
        Else 'Resumen
            lsReporte = lsReporte & reporteOpeInterCMAC_Recepcion_Resumen(mskFechaDe.Text, mskFechaHasta.Text)
        End If
    Else 'Envio
        If optTipo(0).Value Then 'Detallade
            lsReporte = lsReporte & reporteOpeInterCMAC_Envio_Detalle(mskFechaDe.Text, mskFechaHasta.Text)
        Else 'Resumen
            lsReporte = lsReporte & reporteOpeInterCMAC_Envio_Resumen(mskFechaDe.Text, mskFechaHasta.Text)
        End If
    End If
    
    Set P = New Previo.clsPrevio
    Call P.Show(lsReporte, lsTitulo, True)
    Set P = Nothing
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Function reporteOpeInterCMAC_Recepcion_Resumen(pdFechaDe As Date, pdFechaHasta As Date) As String
Dim lsReporte As String, lsCMAC As String, lsSQL As String
Dim lrsReporte As ADODB.Recordset
Dim loConec As New DConecta

    lsSQL = "exec PIT_stp_sel_ReporteOpeInterCMAC_Recepcion_Resumen '" & Format(pdFechaDe, "YYYYMMDD") & "','" & Format(pdFechaHasta, "YYYYMMDD") & "'"
    
    loConec.AbreConexion
    
    Set lrsReporte = loConec.Ejecutar(lsSQL)

    lsReporte = lsReporte & fijarTamanoTexto("CMAC_Origen", 26) & fijarTamanoTexto("Tipo_Operación", 26) & fijarTamanoTexto("Autoriz.", 10, 2) & fijarTamanoTexto("Denegadas", 10, 2) & fijarTamanoTexto("Total", 10, 2) & fijarTamanoTexto("Concil.", 10, 2) & fijarTamanoTexto("Por Conc.", 10, 2) & Chr(10)
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    
    lsCMAC = ""
    
    Do While (Not lrsReporte.EOF)
                
        If lsCMAC <> lrsReporte("cCMACDesc") Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCMACDesc"), 25) & " "
        Else
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        lsCMAC = lrsReporte("cCMACDesc")
        
        If lrsReporte("cOpeDesc") = "TOTAL POR CMAC" And lrsReporte("cCMACDesc") <> "TOTAL GENERAL" Then
            lsReporte = lsReporte & String(74, "-") & Chr(10)
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        If lrsReporte("cCMACDesc") <> "TOTAL GENERAL" Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cOpeDesc"), 25) & " "
        Else
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Autorizadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Denegadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("TotalTX"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Conciliadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("PorConciliar"), 9, 1) & " " & Chr(10)
        
        If lrsReporte("cOpeDesc") = "TOTAL POR CMAC" Then
            lsReporte = lsReporte & Chr(10)
        End If
        
        lrsReporte.MoveNext
    Loop
    
    reporteOpeInterCMAC_Recepcion_Resumen = lsReporte
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function



Public Function reporteOpeInterCMAC_Recepcion_Detalle(pdFechaDe As Date, pdFechaHasta As Date) As String
Dim lsReporte As String, lsCMAC As String, lsSQL As String
Dim lnContAutCMAC As Long, lnContDenCMAC As Long, lnContAutTotal As Long, lnContDenTotal As Long
Dim lnContConcCMAC As Long, lnContPorConcCMAC As Long, lnContConcTotal As Long, lnContPorConcTotal As Long
Dim lnEstAutTX As Integer
Dim lrsReporte As ADODB.Recordset
Dim loConec As New DConecta

    lsSQL = "exec PIT_stp_sel_ReporteOpeInterCMAC_Recepcion_Detalle '" & Format(pdFechaDe, "YYYYMMDD") & "','" & Format(pdFechaHasta, "YYYYMMDD") & "'"
    
    loConec.AbreConexion
    
    Set lrsReporte = loConec.Ejecutar(lsSQL)

    lsReporte = lsReporte & fijarTamanoTexto("CMAC_Origen", 15)
    lsReporte = lsReporte & fijarTamanoTexto("Tipo_Operación", 17)
    lsReporte = lsReporte & fijarTamanoTexto("Fecha", 16, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Tarjeta/DNI", 17, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Cuenta", 19, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Monto", 10, 0)
    lsReporte = lsReporte & fijarTamanoTexto("Moneda", 7, 0)
    lsReporte = lsReporte & fijarTamanoTexto("Estado", 7, 2) & Chr(10)
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    
    lsCMAC = ""
    lnContAutCMAC = 0
    lnContDenCMAC = 0
    lnContAutTotal = 0
    lnContDenTotal = 0
    lnContConcCMAC = 0
    lnContPorConcCMAC = 0
    lnContConcTotal = 0
    lnContPorConcTotal = 0
    lnEstAutTX = -1
    
    Do While (Not lrsReporte.EOF)
        lnContConcTotal = lnContConcTotal + IIf(lrsReporte("nEstado") = 1, 1, 0)
        lnContPorConcTotal = lnContPorConcTotal + IIf(lrsReporte("nEstado") = 0, 1, 0)
        lnContAutTotal = lnContAutTotal + IIf(lrsReporte("nDenegada") = 0, 1, 0)
        lnContDenTotal = lnContDenTotal + IIf(lrsReporte("nDenegada") = 1, 1, 0)
        
        
        If lsCMAC <> lrsReporte("cCMACDesc") Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCMACDesc"), 30) & " "
            lsReporte = lsReporte & Chr(10)
            lsReporte = lsReporte & String(3, " ") & " "
        Else
            lsReporte = lsReporte & String(3, " ") & " "
        End If
        lsCMAC = lrsReporte("cCMACDesc")
        
        If lrsReporte("nDenegada") <> lnEstAutTX Then
            If lrsReporte("nDenegada") = 0 Then 'Autorizada
                lsReporte = lsReporte & fijarTamanoTexto("OPERACIONES AUTORIZADAS", 30, , "*") & " "
            Else
                lsReporte = lsReporte & fijarTamanoTexto("OPERACIONES DENEGADAS", 30, , "*") & " "
            End If
            lsReporte = lsReporte & Chr(10)
            lsReporte = lsReporte & String(3, " ") & " "
            lnEstAutTX = lrsReporte("nDenegada")
        End If
        
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cOpeDesc"), 25) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cFecha"), 15) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cPANTX"), 16) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCuentaTX"), 18) & " "
        lsReporte = lsReporte & fijarTamanoTexto(Format(lrsReporte("nMontoTran"), "##,##0.00"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("nMoneda"), 6, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("nEstado"), 6, 1) & " " & Chr(10)
        
        
        lrsReporte.MoveNext
    Loop
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("RESUMEN          :", 15) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL AUTORIZADAS:", 15) & fijarTamanoTexto(CStr(lnContAutTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL DENEGADAS:", 15) & fijarTamanoTexto(CStr(lnContDenTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL CONCILIADAS:", 15) & fijarTamanoTexto(CStr(lnContConcTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL POR CONCILIAR:", 15) & fijarTamanoTexto(CStr(lnContPorConcTotal), 10, 1) & Chr(10)
    
    
    reporteOpeInterCMAC_Recepcion_Detalle = lsReporte
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function

Public Function reporteOpeInterCMAC_Envio_Resumen(pdFechaDe As Date, pdFechaHasta As Date) As String
Dim lsReporte As String, lsCMAC As String, lsSQL As String
Dim lrsReporte As ADODB.Recordset
Dim loConec As New DConecta

    lsSQL = "exec PIT_stp_sel_ReporteOpeInterCMAC_Envio_Resumen '" & Format(pdFechaDe, "YYYYMMDD") & "','" & Format(pdFechaHasta, "YYYYMMDD") & "'"
    
    loConec.AbreConexion
    
    Set lrsReporte = loConec.Ejecutar(lsSQL)

    lsReporte = lsReporte & fijarTamanoTexto("CMAC_Destino", 26) & fijarTamanoTexto("Tipo_Operación", 26) & fijarTamanoTexto("Autoriz.", 10, 2) & fijarTamanoTexto("Denegadas", 10, 2) & fijarTamanoTexto("Total", 10, 2) & fijarTamanoTexto("Concil.", 10, 2) & fijarTamanoTexto("Por Conc.", 10, 2) & Chr(10)
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    
    lsCMAC = ""
    
    Do While (Not lrsReporte.EOF)
                
        If lsCMAC <> lrsReporte("cCMACDesc") Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCMACDesc"), 25) & " "
        Else
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        lsCMAC = lrsReporte("cCMACDesc")
        
        If lrsReporte("cOpeDesc") = "TOTAL POR CMAC" And lrsReporte("cCMACDesc") <> "TOTAL GENERAL" Then
            lsReporte = lsReporte & String(74, "-") & Chr(10)
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        If lrsReporte("cCMACDesc") <> "TOTAL GENERAL" Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cOpeDesc"), 25) & " "
        Else
            lsReporte = lsReporte & String(25, " ") & " "
        End If
        
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Autorizadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Denegadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("TotalTX"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("Conciliadas"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("PorConciliar"), 9, 1) & " " & Chr(10)
        
        If lrsReporte("cOpeDesc") = "TOTAL POR CMAC" Then
            lsReporte = lsReporte & Chr(10)
        End If
        
        lrsReporte.MoveNext
    Loop
    
    reporteOpeInterCMAC_Envio_Resumen = lsReporte
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function

Public Function reporteOpeInterCMAC_Envio_Detalle(pdFechaDe As Date, pdFechaHasta As Date) As String
Dim lsReporte As String, lsCMAC As String, lsSQL As String
Dim lnContAutCMAC As Long, lnContDenCMAC As Long, lnContAutTotal As Long, lnContDenTotal As Long
Dim lnContConcCMAC As Long, lnContPorConcCMAC As Long, lnContConcTotal As Long, lnContPorConcTotal As Long
Dim lrsReporte As ADODB.Recordset
Dim loConec As New DConecta

    lsSQL = "exec PIT_stp_sel_ReporteOpeInterCMAC_Envio_Detalle '" & Format(pdFechaDe, "YYYYMMDD") & "','" & Format(pdFechaHasta, "YYYYMMDD") & "'"
    
    loConec.AbreConexion
    
    Set lrsReporte = loConec.Ejecutar(lsSQL)

    lsReporte = lsReporte & fijarTamanoTexto("CMAC_Destino", 15)
    lsReporte = lsReporte & fijarTamanoTexto("Tipo_Operación", 17)
    lsReporte = lsReporte & fijarTamanoTexto("Fecha", 16, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Tarjeta/DNI", 17, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Cuenta", 19, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Monto", 10, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Moneda", 7, 2)
    lsReporte = lsReporte & fijarTamanoTexto("Estado", 7, 2) & Chr(10)
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    
    lsCMAC = ""
    lnContAutCMAC = 0
    lnContDenCMAC = 0
    lnContAutTotal = 0
    lnContDenTotal = 0
    lnContConcCMAC = 0
    lnContPorConcCMAC = 0
    lnContConcTotal = 0
    lnContPorConcTotal = 0
    
    Do While (Not lrsReporte.EOF)
        lnContConcTotal = lnContConcTotal + IIf(lrsReporte("nEstado") = 1, 1, 0)
        lnContPorConcTotal = lnContPorConcTotal + IIf(lrsReporte("nEstado") = 0, 1, 0)
        lnContAutTotal = lnContAutTotal + IIf(lrsReporte("nDenegada") = 0, 1, 0)
        lnContDenTotal = lnContDenTotal + IIf(lrsReporte("nDenegada") = 1, 1, 0)
        
        
        If lsCMAC <> lrsReporte("cCMACDesc") Then
            lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCMACDesc"), 30) & " "
            lsReporte = lsReporte & Chr(10)
            lsReporte = lsReporte & String(3, " ") & " "
        Else
            lsReporte = lsReporte & String(3, " ") & " "
        End If
        lsCMAC = lrsReporte("cCMACDesc")
        
        
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cOpeDesc"), 25) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cFecha"), 15) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cPANTX"), 16) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("cCuentaTX"), 18) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("nMontoTran"), 9, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("nMoneda"), 6, 1) & " "
        lsReporte = lsReporte & fijarTamanoTexto(lrsReporte("nEstado"), 6, 1) & " " & Chr(10)
        
        
        lrsReporte.MoveNext
    Loop
    lsReporte = lsReporte & String(100, "-") & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("RESUMEN          :", 15) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL AUTORIZADAS:", 15) & fijarTamanoTexto(CStr(lnContAutTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL DENEGADAS:", 15) & fijarTamanoTexto(CStr(lnContDenTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL CONCILIADAS:", 15) & fijarTamanoTexto(CStr(lnContConcTotal), 10, 1) & Chr(10)
    lsReporte = lsReporte & fijarTamanoTexto("TOTAL POR CONCILIAR:", 15) & fijarTamanoTexto(CStr(lnContPorConcTotal), 10, 1) & Chr(10)
    
    
    reporteOpeInterCMAC_Envio_Detalle = lsReporte
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function

