VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmColPRepoTransacDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Recuperaciones : Reportes Transacciones Diarias "
   ClientHeight    =   4740
   ClientLeft      =   1185
   ClientTop       =   2850
   ClientWidth     =   7680
   HelpContextID   =   200
   Icon            =   "frmColPRepoTransacDia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCodCta 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4665
      TabIndex        =   16
      Top             =   4350
      Width           =   330
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   7290
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   7230
      TabIndex        =   9
      Top             =   4350
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColPRepoTransacDia.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraContenedor 
      Height          =   4155
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   90
      Width           =   7380
      Begin VB.CommandButton cmdBoveda 
         Caption         =   "&Bovedas ..."
         Height          =   345
         Left            =   5940
         TabIndex        =   18
         Top             =   3285
         Width           =   1020
      End
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias ..."
         Height          =   345
         Left            =   5940
         TabIndex        =   17
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Frame fraPagina 
         Caption         =   "Página Inicial "
         Height          =   555
         Left            =   5685
         TabIndex        =   13
         Top             =   2250
         Width           =   1545
         Begin VB.TextBox txtPagIni 
            Height          =   285
            Left            =   270
            TabIndex        =   14
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdEstadMes 
         Caption         =   "&Estadistica Mes"
         Height          =   345
         Left            =   5760
         TabIndex        =   12
         Top             =   3690
         Width           =   1485
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresión "
         Height          =   1095
         Left            =   5685
         TabIndex        =   10
         Top             =   1170
         Width           =   1545
         Begin VB.OptionButton optImpresion 
            Caption         =   "Archivo"
            Height          =   270
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   780
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Impresora"
            Height          =   270
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   495
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Pantalla"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Seleccionar "
         Height          =   1050
         Index           =   0
         Left            =   5700
         TabIndex        =   8
         Top             =   135
         Width           =   1530
         Begin VB.CommandButton cmdSeleccion 
            Caption         =   "Ningun&o"
            Height          =   360
            Index           =   1
            Left            =   270
            TabIndex        =   4
            Top             =   630
            Width           =   1035
         End
         Begin VB.CommandButton cmdSeleccion 
            Caption         =   "&Todos"
            Height          =   360
            Index           =   0
            Left            =   270
            TabIndex        =   3
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.ListBox lstListados 
         Height          =   3660
         ItemData        =   "frmColPRepoTransacDia.frx":038A
         Left            =   120
         List            =   "frmColPRepoTransacDia.frx":038C
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   210
         Width           =   5445
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6210
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   240
      TabIndex        =   11
      Top             =   4335
      Visible         =   0   'False
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmColPRepoTransacDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE VENTA DE REMATE.
'Archivo:  frmColPRepoTransacDia.frm
'LAYG   :  15/09/2001.
'Resumen:  Nos permite registrar una venta de contrato en remate
Option Explicit

Dim pCorte As Variant
Dim pPrevioMax As Double
Dim pItemMax As Double
Dim pProvis0 As Double, pProvis9 As Double, pProvis31 As Double, pProvis61 As Double, pProvis121 As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim pTasaInteresVencido As Double
Dim RegCredPrend As New ADODB.Recordset
Dim sSql As String
Dim pbListGene As Boolean
Dim MuestraImpresion As Boolean
Dim vRTFImp As String
Dim vBuffer As String
Dim vNameForm As String
Dim vNomAge As String
Dim vBoveda As String


'Parametros para el formulario
Private Sub CargaParametros()
    Dim RegPro As New ADODB.Recordset
    Dim tmpSql As String
    On Error GoTo ControlError
    dbCmact.CommandTimeout = 80
    pTasaInteresVencido = ReadParametros("10105")
    pPrevioMax = 10000
    pLineasMax = 56
    pItemMax = 50
    pHojaFiMax = 66
    pProvis0 = 0: pProvis9 = 0: pProvis31 = 0: pProvis61 = 0: pProvis121 = 0:
    tmpSql = " SELECT cCodTab, cvalor " & _
      " FROM " & gcCentralCom & "TablaCod WHERE CCODTAB LIKE '78__' " & _
      " Order BY ccodtab "
    RegPro.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    With RegPro
    If (.BOF Or .EOF) Then
        MsgBox "Ingrese Parametros para la Provisión", vbInformation, " Aviso "
        cmdImprimir.Enabled = False
    Else
        Do While Not .EOF
            If !cCodTab = "7801" Then pProvis0 = Round(Val(!cValor) / 100, 3)
            If !cCodTab = "7802" Then pProvis9 = Round(Val(!cValor) / 100, 3)
            If !cCodTab = "7803" Then pProvis31 = Round(Val(!cValor) / 100, 3)
            If !cCodTab = "7804" Then pProvis61 = Round(Val(!cValor) / 100, 3)
            If !cCodTab = "7805" Then pProvis121 = Round(Val(!cValor) / 100, 3)
            .MoveNext
        Loop
    End If
    End With
    RegPro.Close
    Set RegPro = Nothing
    Exit Sub
    
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdAgencia_Click()
frmPigAgencias.Inicio frmListados
frmPigAgencias.Show 1
End Sub

Private Sub cmdBoveda_Click()
frmPigAgenciaBoveda.Inicio frmListados
frmPigAgenciaBoveda.Show 1
End Sub

Private Sub cmdCodCta_Click()
Dim RegCodCta As New ADODB.Recordset
Dim vLineas As Double, vIndice As Double
Dim vCont As Double
Dim vPage As Double
MousePointer = 11
sSql = "SELECT * FROM relconnueantprend "
RegCodCta.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
vLineas = 3:    vIndice = 0
vPage = 1: vCont = 0
prgList.Min = 0
If RegCodCta.RecordCount > 0 Then prgList.Max = RegCodCta.RecordCount
prgList.Visible = True
vRTFImp = "       Relación del código del Contrato Antiguo con el Código Nuevo" & Space(10) & " Página : " & ImpreFormat(vPage, 6, 0) & gPrnSaltoLinea & gPrnSaltoLinea
With RegCodCta
Do While Not .EOF  'cCodNue      cCodAnt
    vLineas = vLineas + 1
    vIndice = vIndice + 1
    vRTFImp = vRTFImp & ImpreFormat(vIndice, 7, 0) & ImpreFormat(!cCodAnt, 8) & ImpreFormat(!cCodNue, 12)
    If (vIndice Mod 4) = 0 Then vRTFImp = vRTFImp & gPrnSaltoLinea
    If vLineas > pLineasMax * 4 And (vIndice Mod 4) = 0 Then
        vPage = vPage + 1
        vRTFImp = vRTFImp & gPrnSaltoPagina
        vRTFImp = vRTFImp & "       Relación del código del Contrato Antiguo con el Código Nuevo" & Space(10) & " Página : " & ImpreFormat(vPage, 6, 0) & gPrnSaltoLinea & gPrnSaltoLinea
        vLineas = 3
    End If
    vCont = vCont + 1
    prgList.Value = vCont
    Me.Caption = "Registro Nro.: " & vCont
    .MoveNext
Loop
End With
prgList.Visible = False
prgList.Value = 0
Me.Caption = vNameForm
MousePointer = 0
rtfImp.Text = vRTFImp
frmPrevio.Previo rtfImp, " Impresiones Generales ", False, pHojaFiMax
End Sub

Private Sub cmdEstadMes_Click()
        rtfImp.Text = ImprimeEstadMes(frmPigAgencias.List1, prgList, vBoveda)
        frmPrevio.Previo rtfImp, " Listado de Estadística Mensual ", True, pHojaFiMax
    
End Sub

'Permite seleccionar la opción a imprimir y además si se visualiza en pantalla
' o va directo a la impresora
Private Sub cmdImprimir_Click()
Dim x As Integer

On Error GoTo ControlError
pPrevioMax = 4000
vRTFImp = ""
vBuffer = ""
pCorte = IIf(optImpresion(0).Value = True, gPrnSaltoLinea, vbCrLf)
MuestraImpresion = False
' Carga Bovedas Seleccionadas
vBoveda = CargaBovedaSelec
'Verifica si está preparada la impresora cuando se envia directo a ella
If optImpresion(1).Value = True Then
    If Not ImpreSensa Then Exit Sub
End If
With lstListados
    If .Selected(0) = True Then  'Préstamos No Desembolsados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresDife", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresDife", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
                    
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If

    If .Selected(1) = True Then  'Prest. Diferidos
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresNorm", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresNorm", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
                    
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(2) = True Then  'Prést. Normales
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresVencNorm", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresVencNorm", dbCmact
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(3) = True Then  'Prést. Vencidos Normales
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresParaRemaNorm", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresParaRemaNorm", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(4) = True Then  'Prest. ParaRemate Normales
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresReno", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresReno", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(5) = True Then  'Prest. Renovados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresVencReno", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresVencReno", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(6) = True Then  'Prest. Vencidos Renovados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresParaRemaReno", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                    ImprimePrestamo "PresParaRemaReno", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(7) = True Then  'Prest. Vigentes (Auditoria)
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresVige", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresVige", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(8) = True Then  'Listado de Intereses Devengados - General
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeIntDev "InteDeve", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeIntDev "InteDeve", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(9) = True Then  'Listado de Intereses Devengados - (1 - 30 )
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeIntDev "InteDeve01", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeIntDev "InteDeve01", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(10) = True Then  'Listado de Intereses Devengados - (31 - 120 )
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeIntDev "InteDeve31", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeIntDev "InteDeve31", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(11) = True Then  'Listado de Intereses Devengados- (121 - ... )
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeIntDev "InteDeve121", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeIntDev "InteDeve121", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(12) = True Then  'Listado de Provisiones
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeProvis "Provis", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeProvis "Provis", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
                    
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    If .Selected(13) = True Then  'Listado de Provisiones  (Consolidada-CreditoAudi)
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimeProvis2 "Provis", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimeProvis2 "Provis", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(14) = True Then  'Prest. Vigentes en Boveda (Sin Diferidos)
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    ImprimePrestamo "PresVigeBove", dbCmact
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        ImprimePrestamo "PresVigeBove", dbCmactN
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        If Not MuestraImpresion Then
            If Len(vBuffer) > 0 Then
                vBuffer = Left(vBuffer, Len(vBuffer) - 1)
                MuestraImpresion = True
            End If
        End If
    End If

    If .Selected(15) = True Then  ' Estad Mensual de Creditos (Distribucion de Creditos)
        'ImprimeEstadMes
        If MuestraImpresion Then vBuffer = vBuffer & gPrnSaltoPagina
        vBuffer = vBuffer & ImprimeEstadMes(frmPigAgencias.List1, prgList, vBoveda)
    End If


End With
'Envia a la impresion Previa
If optImpresion(0).Value = True And Len(Trim(vBuffer)) > 0 Then
    rtfImp.Text = vBuffer
    frmPrevio.Previo rtfImp, " Impresiones Generales ", True, pHojaFiMax
ElseIf optImpresion(2).Value = True And Len(Trim(vBuffer)) > 0 Then
    dlgGrabar.CancelError = True
    dlgGrabar.InitDir = App.Path
    dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
    dlgGrabar.ShowSave
    If dlgGrabar.FileName <> "" Then
        Open dlgGrabar.FileName For Output As #1
        Print #1, vBuffer
        Close #1
    End If
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If
End Sub

'Permite Selecionar o quitar la seleción de las opciones de los listados
Private Sub cmdSeleccion_Click(Index As Integer)
If Index = 0 Then
    OpcListado lstListados, True
Else
    OpcListado lstListados, False
End If
End Sub

'Permite salir del formulario actual
Private Sub cmdSalir_Click()
Unload Me
End Sub

'Permite inicializar el formulario
Private Sub Form_Load()
AbreConexion
Me.Enabled = True
Dim RegUsu As New ADODB.Recordset
gcIntCentra = CentraSdi(Me)
CargaParametros
txtPagIni.Text = "1"
vNameForm = "Crédito Pignoraticio : Listados General"
With lstListados
    .Clear
    .AddItem "Listado de Préstamos Diferidos"
    .AddItem "Listado de Préstamos Normales"
    .AddItem "Listado de Préstamos Vencidos Normales"
    .AddItem "Listado de Préstamos Para Remate Normales"
    .AddItem "Listado de Préstamos Renovados"
    .AddItem "Listado de Préstamos Vencidos Renovados"
    .AddItem "Listado de Préstamos Para Remate Renovados"
    .AddItem "Listado de Préstamos General - Vigentes y Cancelados"
    .AddItem "Listado de Crédito Pignoraticio (todos los rangos) "
    .AddItem "Listado de Crédito Pignoraticio (hasta 30 dias)"
    .AddItem "Listado de Crédito Pignoraticio (31 a 120 dias)"
    .AddItem "Listado de Crédito Pignoraticio (121 a más dias)"
    .AddItem "Listado de Provisiones - Créditos al cierre de mes entre dias de mora"
    .AddItem "Listado de Provisiones Consolidada - Al cierre de mes por Calificaciones"
    .AddItem "Listado de Préstamos Vigentes en Boveda (Sin Diferidos) "
    .AddItem "Estadistica Mensual (Dist. Creditos)"
End With
End Sub

'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer, Optional ByVal pPagCorta As Boolean = True)
    Dim vTitulo As String
    Dim vSubTit As String
    'Dim vSpaTit As Integer
    'Dim vSpaSub As Integer
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    vSubTit = ""
    Select Case vOpt
        Case "PresDife"
            vTitulo = "LISTADO DE PRESTAMOS DIFERIDOS"
        Case "PresNorm"
            vTitulo = "LISTADO DE PRESTAMOS NORMALES"
        Case "PresVencNorm"
            vTitulo = "LISTADO DE PRESTAMOS VENCIDOS NORMALES"
        Case "PresParaRemaNorm"
            vTitulo = "LISTADO DE PRESTAMOS PARA REMATE NORMALES"
        Case "PresReno"
            vTitulo = "LISTADO DE PRESTAMOS RENOVADOS"
        Case "PresVencReno"
            vTitulo = "LISTADO DE PRESTAMOS VENCIDOS RENOVADOS"
        Case "PresParaRemaReno"
            vTitulo = "LISTADO DE PRESTAMOS PARA REMATE RENOVADOS"
        Case "PresVige"
            vTitulo = "LISTADO DE PRESTAMOS VIGENTES"
        Case "InteDeve"
            vTitulo = "LISTADO DE CREDITO PIGNORATICIO"
            vSubTit = " GENERAL "
        Case "InteDeve01"
            vTitulo = "LISTADO DE CREDITO PIGNORATICIO"
            vSubTit = " 0 - 30 DIAS "
        Case "InteDeve31"
            vTitulo = "LISTADO DE CREDITO PIGNORATICIO"
            vSubTit = " 31 - 120 DIAS "
        Case "InteDeve121"
            vTitulo = "LISTADO DE CREDITO PIGNORATICIO"
            vSubTit = " 121 DIAS A MAS "
        Case "Provis"
            vTitulo = "LISTADO DE PROVISIONES AL CIERRE DE MES ENTRE DIAS DE MORA"
        Case "Provis2"
            vTitulo = "LISTADO DE PROVISIONES CONSOLIDADAS AL CIERRE DE MES POR CALIFICACION"
        Case "PresVigeBove"
            vTitulo = "LISTADO DE PRESTAMOS VIGENTES BOVEDA (SIN DIFERIDOS)"
            
    End Select
    vArea = "Crédito Pignoraticio"
    vNroLineas = IIf(pPagCorta = True, 140, 159)
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2) - 25, " ") & vTitulo & String(Round((vNroLineas - Len(Trim(vTitulo))) / 2) - 29, " ")
    'Centra SubTítulo
    vSubTit = String(Round((vNroLineas - Len(Trim(vSubTit))) / 2) - 25, " ") & vSubTit & String(Round((vNroLineas - Len(Trim(vSubTit))) / 2) - 29, " ")
    
    vRTFImp = vRTFImp & pCorte
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(vNomAge, 25, 0) & vTitulo & Space(11) & "Página: " & Format(vPagina, "@@@@") & pCorte
    vRTFImp = vRTFImp & Space(1) & vArea & vSubTit & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & pCorte
    vRTFImp = vRTFImp & String(vNroLineas, "-") & pCorte
    If Left(vOpt, 8) <> "InteDeve" And vOpt <> "Provis" And vOpt <> "Provis2" Then
        vRTFImp = vRTFImp & Space(2) & " ITEM      CONTRATO     COD.ANT.   APELLIDOS Y NOMBRES           F.EMPEÑO PLAZ.   PRESTAMO       SALDO      TASACION ORO NETO EST BOVE BLOQ" & pCorte
    ElseIf (vOpt = "InteDeve" Or vOpt = "InteDeve01" Or vOpt = "InteDeve31" Or vOpt = "InteDeve121") Then
        'F.VENCIM.   M.ACT. M.PROX.
        vRTFImp = vRTFImp & Space(2) & " ITEM    CONTRATO    F.EMPEÑO     APELLIDOS Y NOMBRES         SALDO     TASACION  PL. INT.DEV. ATRASO INT.M.ACT. INT.M.PROX. TOT.INT.COB.  DEU.TOTAL ORO NETO" & pCorte
    ElseIf vOpt = "Provis" Or vOpt = "Provis2" Then
        vRTFImp = vRTFImp & Space(2) & " ITEM    CONTRATO    F.EMPEÑO     APELLIDOS Y NOMBRES          SALDO     TASACION   PL.  F.VENCIM.  INT.DEV. D.ATRA. DEUDA TOTAL  ORO NETO   PORC.    PROVIS." & pCorte
    End If
    vRTFImp = vRTFImp & String(vNroLineas, "-") & pCorte
End Sub

'Imprime Diferentes Préstamos de acuerdo a condición
Private Sub ImprimePrestamo(ByVal CondicionImpresion As String, pConexPrend As ADODB.Connection)
    Dim vNombre As String * 32
    Dim vIndice As Long  'contador de Item
    Dim vLineas As Integer
    Dim vPage As Integer
    Dim vTotPres As Currency, vTotVaTa As Currency, vTotOrNe As Currency, vTotSald As Currency
    Dim vCabecera As String
    Dim vCont As Long
    Dim lsBloq As String * 5
    MousePointer = 11
    MuestraImpresion = True
    sSql = "SELECT cp.cCodCta, cp.dfecpres, cp.nplazo, cp.nprestamo, cp.nsaldocap, cp.cestado, " & _
        " cp.nValTasac, cp.nOroNeto, cp.nOroBruto, r.cCodAnt, cp.cAgeBoveda " & _
        " FROM CredPrenda CP LEFT JOIN RelConNueAntPrend R ON cp.ccodcta = r.ccodnue "
    Select Case CondicionImpresion
        Case "PresDife"
            sSql = sSql & " WHERE cp.cestado = '2' "
            vCabecera = "PresDife"
        Case "PresNorm"
            sSql = sSql & " WHERE cp.cestado = '1' AND cp.nnumrenov = 0  "
            vCabecera = "PresNorm"
        Case "PresVencNorm"
            sSql = sSql & " WHERE cp.cestado = '4' AND cp.nnumrenov = 0 "
            vCabecera = "PresVencNorm"
        Case "PresParaRemaNorm"
            sSql = sSql & " WHERE cp.cestado = '6' AND cp.nnumrenov = 0 "
            vCabecera = "PresParaRemaNorm"
        Case "PresReno"
            sSql = sSql & " WHERE cp.cestado = '7' AND cp.nnumrenov > 0 "
            vCabecera = "PresReno"
        Case "PresVencReno"
            sSql = sSql & " WHERE cp.cestado = '4' AND cp.nnumrenov > 0 "
            vCabecera = "PresVencReno"
        Case "PresParaRemaReno"
            sSql = sSql & " WHERE cp.cestado = '6' AND cp.nnumrenov > 0 "
            vCabecera = "PresParaRemaReno"
        Case "PresVige"
            sSql = sSql & " WHERE cp.cestado IN ('1','2','4','6','7') "
            vCabecera = "PresVige"
        Case "PresVigeBove"  ' Sin Diferidos
            sSql = sSql & " WHERE cp.cestado IN ('1','4','6','7') "
            vCabecera = "PresVigeBove"
        Case Else
            MsgBox "Opción no reconocida", vbInformation, " ! Aviso ! "
            Exit Sub
    End Select
    If Len(Trim(vBoveda)) > 0 Then
        sSql = sSql & " AND cp.cAgeBoveda in " & vBoveda & " "
    End If
    sSql = sSql & " ORDER BY cp.cCodCta"
    RegCredPrend.Open sSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        Select Case CondicionImpresion
            Case "PresDife"
                MsgBox " No existen Préstamos Diferidos, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresNorm"
                MsgBox " No existen Préstamos Normales, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresVencNorm"
                MsgBox " No existen Préstamos Vencidos Normales, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresParaRemaNorm"
                MsgBox " No existen Préstamos para Remate Normales, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresReno"
                MsgBox " No existen Préstamos Renovados,  Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresVencReno"
                MsgBox " No existen Préstamos Vencidos Renovados, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresParaRemaReno"
                MsgBox " No existen Préstamos para Remate Renovados, Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresVige"
                MsgBox " No existen Préstamos Vigentes,  Bovedas " & vBoveda, vbInformation, " ! Aviso ! "
            Case "PresVigeBove"
                MsgBox " No existen Préstamos Vigentes en Boveda (Sin Diferidos)", vbInformation, " ! Aviso ! "
        End Select
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegCredPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegCredPrend.Close
                Set RegCredPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage
            End If
        ElseIf optImpresion(1).Value = True Then
            ImpreBegin True, pHojaFiMax
            vRTFImp = ""
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage
                Print #ArcSal, ImpreCarEsp(vRTFImp);
                vRTFImp = ""
            End If
        Else
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage
            End If
        End If
        prgList.Visible = True
        vTotPres = 0: vTotVaTa = 0: vTotOrNe = 0
        vIndice = 1
        vLineas = 7
        With RegCredPrend
            Do While Not .EOF
                vNombre = PstaNombre(ClienteNombre(!cCodCta, pConexPrend), False)
                If CondicionImpresion = "PresDife" Then
                   lsBloq = IIf(IsCtaBlo(!cCodCta, pConexPrend), "BLOQ", "  ")
                Else
                   lsBloq = "    "
                End If
                If vIndice > ((Val(txtPagIni) * pItemMax) - pItemMax) Then
                If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        ImpreFormat("" & !cCodAnt, 8, 1) & ImpreFormat(vNombre, 30, 1) & _
                        Space(1) & Format(!dfecpres, "dd/mm/yyyy") & _
                        ImpreFormat(!nPlazo, 3, 0) & ImpreFormat(!nPrestamo, 10) & ImpreFormat(!nSaldoCap, 10) & _
                        ImpreFormat(!nvaltasac, 10) & ImpreFormat(!noroneto, 7) & ImpreFormat(!cEstado, 2) & Space(1) & Mid(!cAgeBoveda, 4, 2) & lsBloq & pCorte
                    If vIndice Mod 300 = 0 Then
                        vBuffer = vBuffer & vRTFImp
                        vRTFImp = ""
                    End If
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        ImpreFormat("" & !cCodAnt, 8, 1) & ImpreCarEsp(ImpreFormat(vNombre, 30, 1)) & _
                        Space(1) & Format(!dfecpres, "dd/mm/yyyy") & _
                        ImpreFormat(!nPlazo, 3, 0) & ImpreFormat(!nPrestamo, 10) & ImpreFormat(!nSaldoCap, 10) & _
                        ImpreFormat(!nvaltasac, 10) & ImpreFormat(!noroneto, 7) & ImpreFormat(!cEstado, 2) & Space(1) & Mid(!cAgeBoveda, 4, 2) & lsBloq
                End If
                End If
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                vTotPres = vTotPres + !nPrestamo
                vTotVaTa = vTotVaTa + !nvaltasac
                vTotOrNe = vTotOrNe + !noroneto
                vTotSald = vTotSald + !nSaldoCap
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If vPage > (Val(txtPagIni) - 1) Then
                    If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                        If Val(txtPagIni) <> vPage Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        End If
                        Cabecera vCabecera, vPage
                    Else
                        If Val(txtPagIni) <> vPage Then
                            If vPage Mod 5 = 0 Then
                                ImpreEnd
                                ImpreBegin True, pHojaFiMax
                            Else
                                ImpreNewPage
                            End If
                        End If
                        vRTFImp = ""
                        Cabecera vCabecera, vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If Val(txtPagIni) <= vPage Then
            If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & ImpreFormat("TOTAL", 6, 27) & ImpreFormat((vIndice - 1), 6, 0) & _
                    ImpreFormat(vTotPres, 50) & ImpreFormat(vTotSald, 10) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(Round(vTotOrNe, 2), 7) & pCorte
                vBuffer = vBuffer & vRTFImp
                vRTFImp = ""
            Else
                Print #ArcSal, " "
                Print #ArcSal, ImpreFormat("TOTAL", 6, 27) & ImpreFormat((vIndice - 1), 6, 0) & _
                    ImpreFormat(vTotPres, 50) & ImpreFormat(vTotSald, 10) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(Round(vTotOrNe, 2), 7)
                ImpreEnd
            End If
            End If
        End With
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Imprime Listado para Intereses Devengados (para Auditoria)
Private Sub ImprimeIntDev(ByVal CondicionImpresion As String, pConexPrend As ADODB.Connection)
Dim vAtraso As Double, vDiaAct As Double, vDiaPro As Double
Dim vIntDev As Currency, vIntAct As Currency, vIntPro As Currency
Dim vIntAde As Currency
Dim vFecVen As String
Dim vNombre As String * 35
Dim vIndice As Double  'contador de Item
Dim vLineas As Double
Dim vPage As Double
Dim vTotSald As Currency, vTotVaTa As Currency
Dim vTotIntDev As Currency
Dim vTotDAtra As Currency, vTotDAct As Currency, vTotDPro As Currency
Dim vTotIAct As Currency, vTotIPro As Currency
Dim vTotOrNe As Currency
Dim vCabecera As String
Dim vCont As Double
    MousePointer = 11
    MuestraImpresion = True
    sSql = "SELECT cp.cCodCta, cp.dfecpres, cp.nplazo, cp.nprestamo, cp.nsaldocap, " & _
        " cp.nValTasac, cp.nOroNeto, cp.dfecvenc, cp.ntasaint " & _
        " FROM CredPrenda CP "
    Select Case CondicionImpresion
        Case "InteDeve"
            sSql = sSql & " WHERE cp.cestado IN ('1','4','7','6') "
            vCabecera = "InteDeve"
        Case "InteDeve01"
            sSql = sSql & " WHERE cp.cestado IN ('1','4','7','6') AND " & _
             " datediff(dd, cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') < 31 "
            vCabecera = "InteDeve01"
        Case "InteDeve31"
            sSql = sSql & " WHERE cp.cestado IN ('1','4','7','6') AND " & _
             " datediff(dd, cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') > 30  AND " & _
             " datediff(dd, cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') < 121 "
            vCabecera = "InteDeve31"
        Case "InteDeve121"
            sSql = sSql & " WHERE cp.cestado IN ('1','4','7','6') AND " & _
             " datediff(dd, cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') > 120 "
            vCabecera = "InteDeve121"
        Case Else
            MsgBox "Opción no reconocida", vbInformation, " ! Aviso ! "
            Exit Sub
    End Select
    If Len(Trim(vBoveda)) > 0 Then
        sSql = sSql & " AND cp.cAgeBoveda in " & vBoveda & " "
    End If
    sSql = sSql & " ORDER BY  cp.ccodcta"
    RegCredPrend.CursorLocation = adUseClient
    RegCredPrend.Open sSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
    Set RegCredPrend.ActiveConnection = Nothing
    
    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        Select Case CondicionImpresion
            Case "InteDeve", "InteDeve01", "InteDeve31", "InteDeve121"
                MsgBox " No existen Contratos ", vbInformation, " ! Aviso ! "
        End Select
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegCredPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegCredPrend.Close
                Set RegCredPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        ElseIf optImpresion(1).Value = True Then
            ImpreBegin True, pHojaFiMax
            vRTFImp = ""
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
                Print #ArcSal, ImpreCarEsp(vRTFImp);
                vRTFImp = ""
            End If
        Else
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        End If
        prgList.Visible = True
        vTotVaTa = 0: vTotOrNe = 0
        vIndice = 1
        vLineas = 7
        With RegCredPrend
            Do While Not .EOF
                vIntDev = 0:    vIntAct = 0:    vIntPro = 0:    vIntAde = 0
                vAtraso = 0:    vDiaAct = 0:    vDiaPro = 0
                vFecVen = Format(!dFecVenc, "dd/mm/yyyy")
                vNombre = PstaNombre(ClienteNombre(!cCodCta, pConexPrend), False)
                If DateDiff("d", vFecVen, gdFecSis) > 0 Then
                    vIntDev = ((1 + pTasaInteresVencido) ^ (DateDiff("d", vFecVen, gdFecSis) / 30) - 1) * !nSaldoCap
                    vIntDev = Round(vIntDev, 2)
                End If
                vAtraso = IIf(DateDiff("d", vFecVen, gdFecSis) > 0, DateDiff("d", vFecVen, gdFecSis), 0)
                If vAtraso = 0 Then
                    vDiaPro = IIf(DateDiff("d", gdFecSis, vFecVen) > 0, DateDiff("d", gdFecSis, vFecVen), 0)
                    vDiaAct = !nPlazo - vDiaPro
                    vIntAde = CalculaInteresAdelantado(!nSaldoCap, !nTasaInt, !nPlazo)
                    'vIntAde = Round(vIntAde, 2)
                    vIntAct = !nSaldoCap * ((((vIntAde / !nSaldoCap + 1) ^ (1 / !nPlazo)) ^ vDiaAct) - 1)
                    vIntAct = Round(vIntAct, 2)
                    vIntPro = (!nSaldoCap + vIntAct) * ((((vIntAde / !nSaldoCap + 1) ^ (1 / !nPlazo)) ^ vDiaPro) - 1)
                    vIntPro = Round(vIntPro, 2)
                End If
                If vIndice > ((Val(txtPagIni) * pItemMax) - pItemMax) Then
                If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreFormat(vNombre, 21, 1) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 3, 0) & _
                        ImpreFormat(vIntDev, 6) & ImpreFormat(vAtraso, 5, 0) & _
                        ImpreFormat(vIntAct, 9) & ImpreFormat(vIntPro, 9) & _
                        ImpreFormat(vIntAct + vIntPro, 9) & ImpreFormat(!nSaldoCap + vIntDev, 10) & _
                        ImpreFormat(!noroneto, 7) & pCorte
                        '& Space(1) & Format(!dFecVenc, "dd/mm/yyyy")
                        'ImpreFormat(vDiaAct, 6, 0) & ImpreFormat(vDiaPro, 6, 0) &
                    If vIndice Mod 300 = 0 Then
                        vBuffer = vBuffer & vRTFImp
                        vRTFImp = ""
                    End If
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreCarEsp(ImpreFormat(vNombre, 21, 1)) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 3, 0) & _
                        ImpreFormat(vIntDev, 6) & ImpreFormat(vAtraso, 5, 0) & _
                        ImpreFormat(vIntAct, 9) & ImpreFormat(vIntPro, 9) & _
                        ImpreFormat(vIntAct + vIntPro, 9) & ImpreFormat(!nSaldoCap + vIntDev, 10) & _
                        ImpreFormat(!noroneto, 7)
                        'Space(1) & Format(!dFecVenc, "dd/mm/yyyy") &
                        'ImpreFormat(vDiaAct, 6, 0) & ImpreFormat(vDiaPro, 6, 0) &
                End If
                End If
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                vTotSald = vTotSald + !nSaldoCap
                vTotVaTa = vTotVaTa + !nvaltasac
                vTotIntDev = vTotIntDev + vIntDev
                vTotDAtra = vTotDAtra + vAtraso
                vTotDAct = vTotDAct + vDiaAct
                vTotDPro = vTotDPro + vDiaPro
                vTotIAct = vTotIAct + vIntAct
                vTotIPro = vTotIPro + vIntPro
                vTotOrNe = vTotOrNe + !noroneto
                If vLineas > pLineasMax Then
                    DoEvents
                    vPage = vPage + 1
                    If vPage > (Val(txtPagIni) - 1) Then
                    If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                        If Val(txtPagIni) <> vPage Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        End If
                        Cabecera vCabecera, vPage, False
                    Else
                        If Val(txtPagIni) <> vPage Then
                            If vPage Mod 5 = 0 Then
                                ImpreEnd
                                ImpreBegin True, pHojaFiMax
                            Else
                                ImpreNewPage
                            End If
                        End If
                        vRTFImp = ""
                        Cabecera vCabecera, vPage, False
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If Val(txtPagIni) <= vPage Then
            If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 39) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 9) & _
 _
                    ImpreFormat(vTotIAct, 14) & ImpreFormat(vTotIPro, 9) & _
                    ImpreFormat(vTotIAct + vTotIPro, 9) & ImpreFormat(vTotSald + vTotIntDev, 10) & _
                    ImpreFormat(Round(vTotOrNe, 2), 7) & pCorte
                    'ImpreFormat(vTotDAtra, 7, 0) &
                    'ImpreFormat(vTotDAct, 7, 0) & ImpreFormat(vTotDPro, 7, 0) &
                vBuffer = vBuffer & vRTFImp
                vRTFImp = ""
            Else
                Print #ArcSal, " "
                Print #ArcSal, ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 39) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 9) & _
 _
                    ImpreFormat(vTotIAct, 14) & ImpreFormat(vTotIPro, 9) & _
                    ImpreFormat(vTotIAct + vTotIPro, 9) & ImpreFormat(vTotSald + vTotIntDev, 10) & _
                    ImpreFormat(Round(vTotOrNe, 2), 7)
                    'ImpreFormat(vTotDAtra, 7, 0) &
                    'ImpreFormat(vTotDAct, 7, 0) & ImpreFormat(vTotDPro, 7, 0) &
                ImpreEnd
            End If
            End If
        End With
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Imprime Listado para Intereses Devengados (para Auditoria)
Private Sub ImprimeProvis(ByVal CondicionImpresion As String, pConexPrend As ADODB.Connection)
Dim vAtraso As Integer
Dim vPorcen As Double, vProvis  As Currency
Dim vFecVen As String
Dim vIntDev As Currency
Dim vNombre As String * 35
Dim vIndice As Double  'contador de Item
Dim vLineas As Integer
Dim vPage As Integer
Dim vTotSald As Currency, vTotVaTa As Currency
Dim vTotIntDev As Currency
Dim vTotOrNe As Currency, vTotProvis As Currency
Dim vCabecera As String
Dim vCont As Double
Dim vInd0 As Double, vSal0 As Currency
Dim vInd9 As Double, vSal9 As Currency
Dim vInd31 As Double, vSal31 As Currency
Dim vInd61 As Double, vSal61 As Currency
Dim vInd121 As Double, vSal121 As Currency
Dim v0 As Boolean, v9 As Boolean, v31 As Boolean, v61 As Boolean, v121 As Boolean
    MousePointer = 11
    MuestraImpresion = True
    sSql = "SELECT cp.cCodCta, cp.dfecpres, cp.nplazo, cp.nprestamo, cp.nsaldocap, " & _
        " cp.nValTasac, cp.nOroNeto, cp.dfecvenc, cp.ntasaint, datediff(dd,cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') AS DiasDife " & _
        " FROM CredPrenda CP " & _
        " WHERE cp.cestado IN ('1','4','7','6') " & _
        " ORDER BY DiasDife, cp.dfecpres, cp.cCodCta "
    vCabecera = "Provis"
    RegCredPrend.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        Select Case CondicionImpresion
            Case "Provis"
                MsgBox " No existen Contratos ", vbInformation, " ! Aviso ! "
        End Select
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegCredPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegCredPrend.Close
                Set RegCredPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        ElseIf optImpresion(1).Value = True Then
            ImpreBegin True, pHojaFiMax
            vRTFImp = ""
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
                Print #ArcSal, ImpreCarEsp(vRTFImp);
                vRTFImp = ""
            End If
        Else
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        End If
        prgList.Visible = True
        vPorcen = 0: vProvis = 0
        vTotVaTa = 0: vTotOrNe = 0: vTotProvis = 0
        v0 = True: v9 = True: v31 = True: v61 = True: v121 = True
        vInd0 = 0: vInd9 = 0: vInd31 = 0: vInd61 = 0: vInd121 = 0:
        vSal0 = 0: vSal9 = 0: vSal31 = 0: vSal61 = 0: vSal121 = 0:
        vIndice = 1
        vLineas = 7
        With RegCredPrend
            Do While Not .EOF
                If !diasdife <= 8 Then
                    If v0 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DIAS DE MORA DE : 0 A 8 DIAS" & pCorte
                        End If
                        v0 = False
                        vPorcen = pProvis0
                    End If
                    vInd0 = vInd0 + 1: vSal0 = vSal0 + !nSaldoCap
                ElseIf !diasdife >= 9 And !diasdife <= 30 Then
                    If v9 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DIAS DE MORA DE : 9 A 30 DIAS" & pCorte
                        End If
                        v9 = False
                        vPorcen = pProvis9
                    End If
                    vInd9 = vInd9 + 1: vSal9 = vSal9 + !nSaldoCap
                ElseIf !diasdife >= 31 And !diasdife <= 60 Then
                    If v31 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DIAS DE MORA DE : 31 A 60 DIAS" & pCorte
                        End If
                        v31 = False
                        vPorcen = pProvis31
                    End If
                    vInd31 = vInd31 + 1: vSal31 = vSal31 + !nSaldoCap
                ElseIf !diasdife >= 61 And !diasdife <= 120 Then
                    If v61 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DIAS DE MORA DE : 61 A 120 DIAS" & pCorte
                        End If
                        v61 = False
                        vPorcen = pProvis61
                    End If
                    vInd61 = vInd61 + 1: vSal61 = vSal61 + !nSaldoCap
                ElseIf !diasdife >= 121 Then
                    If v121 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DIAS DE MORA DE : MAS DE 120 DIAS" & pCorte
                        End If
                        v121 = False
                        vPorcen = pProvis121
                    End If
                    vInd121 = vInd121 + 1: vSal121 = vSal121 + !nSaldoCap
                End If
                vAtraso = 0
                vFecVen = Format(!dFecVenc, "dd/mm/yyyy")
                vNombre = PstaNombre(ClienteNombre(!cCodCta, dbCmact), False)
                If DateDiff("d", vFecVen, gdFecSis) > 0 Then
                    vIntDev = ((1 + pTasaInteresVencido) ^ (DateDiff("d", vFecVen, gdFecSis) / 30) - 1) * !nSaldoCap
                    vIntDev = Round(vIntDev, 2)
                End If
                vAtraso = IIf(DateDiff("d", vFecVen, gdFecSis) > 0, DateDiff("d", vFecVen, gdFecSis), 0)
                vProvis = !nSaldoCap * vPorcen
                vProvis = Round(vProvis, 2)
                If vIndice > ((Val(txtPagIni) * pItemMax) - pItemMax) Then
                If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 8, 0) & ImpreFormat(!cCodCta, 13, 1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreFormat(vNombre, 25, 1) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 4, 0) & Space(1) & Format(!dFecVenc, "dd/mm/yyyy") & _
                        ImpreFormat(vIntDev, 7) & ImpreFormat(vAtraso, 6, 0) & _
                        ImpreFormat(!nSaldoCap + vIntDev, 10) & ImpreFormat(!noroneto, 7) & _
                        ImpreFormat(vPorcen, 5, 3) & ImpreFormat(vProvis, 8) & pCorte
                    If vIndice Mod 300 = 0 Then
                        vBuffer = vBuffer & vRTFImp
                        vRTFImp = ""
                    End If
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 8, 0) & ImpreFormat(!cCodCta, 13, 1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreCarEsp(ImpreFormat(vNombre, 25, 1)) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 4, 0) & Space(1) & Format(!dFecVenc, "dd/mm/yyyy") & _
                        ImpreFormat(vIntDev, 7) & ImpreFormat(vAtraso, 6, 0) & _
                        ImpreFormat(!nSaldoCap + vIntDev, 10) & ImpreFormat(!noroneto, 7) & _
                        ImpreFormat(vPorcen, 5, 3) & ImpreFormat(vProvis, 8)
                End If
                End If
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                vTotSald = vTotSald + !nSaldoCap
                vTotVaTa = vTotVaTa + !nvaltasac
                vTotIntDev = vTotIntDev + vIntDev
                vTotOrNe = vTotOrNe + !noroneto
                vTotProvis = vTotProvis + vProvis
                
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If vPage > (Val(txtPagIni) - 1) Then
                    If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                        If Val(txtPagIni) <> vPage Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        End If
                        Cabecera vCabecera, vPage, False
                    Else
                        If Val(txtPagIni) <> vPage Then
                            If vPage Mod 5 = 0 Then
                                ImpreEnd
                                ImpreBegin True, pHojaFiMax
                            Else
                                ImpreNewPage
                            End If
                        End If
                        vRTFImp = ""
                        Cabecera vCabecera, vPage, False
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotSald = Round(vTotSald, 2)
            If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                If Val(txtPagIni) <= vPage Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 40) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 22) & _
                    ImpreFormat(vTotSald + vTotIntDev, 16) & ImpreFormat(Round(vTotOrNe, 2), 7) & _
                    ImpreFormat(Round(vTotProvis, 2), 17) & pCorte
                vRTFImp = vRTFImp & gPrnSaltoPagina
                End If
                vPage = vPage + 1
                If Val(txtPagIni) <= vPage Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(83) & "Página: " & Format(vPage, "@@@@") & pCorte
                vRTFImp = vRTFImp & Space(1) & "Crédito Pignoraticio" & Space(81) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & "       CONSOLIDADO DE CREDITOS AL CIERRE DE MES ENTRE DIAS DE MORA " & pCorte
                vRTFImp = vRTFImp & Space(25) & "       =========================================================== " & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & "   TOTAL    CLASIFICACION      SALDO CAPITAL      PORCEN.    PROVISION" & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd0, 7, 0) & "   Normal            " & ImpreFormat(vSal0, 12) & ImpreFormat(pProvis0, 10, 3) & ImpreFormat(vSal0 * pProvis0, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd9, 7, 0) & "   C.P.P.            " & ImpreFormat(vSal9, 12) & ImpreFormat(pProvis9, 10, 3) & ImpreFormat(vSal9 * pProvis9, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd31, 7, 0) & "   Deficiente        " & ImpreFormat(vSal31, 12) & ImpreFormat(pProvis31, 10, 3) & ImpreFormat(vSal31 * pProvis31, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd61, 7, 0) & "   Dudoso            " & ImpreFormat(vSal61, 12) & ImpreFormat(pProvis61, 10, 3) & ImpreFormat(vSal61 * pProvis61, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd121, 7, 0) & "   Perdida           " & ImpreFormat(vSal121, 12) & ImpreFormat(pProvis121, 10, 3) & ImpreFormat(vSal121 * pProvis121, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(10) & "TOTAL" & Space(10) & ImpreFormat(vInd0 + vInd9 + vInd31 + vInd61 + vInd121, 7, 0) & _
                    Space(21) & ImpreFormat(vSal0 + vSal9 + vSal31 + vSal61 + vSal121, 12) & _
                    Space(14) & ImpreFormat(Round((vSal0 * pProvis0) + (vSal9 * pProvis9) + (vSal31 * pProvis31) + (vSal61 * pProvis61) + (vSal121 * pProvis121), 2), 10, 2) & pCorte
                vRTFImp = vRTFImp & pCorte
                End If
                vBuffer = vBuffer & vRTFImp
                vRTFImp = ""
            Else
                If Val(txtPagIni) <= vPage Then
                Print #ArcSal, " "
                Print #ArcSal, ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 40) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 22) & _
                    ImpreFormat(vTotSald + vTotIntDev, 16) & ImpreFormat(Round(vTotOrNe, 2), 7) & _
                    ImpreFormat(vTotProvis, 17)
                ImpreNewPage
                End If
                vPage = vPage + 1
                If Val(txtPagIni) <= vPage Then
                Print #ArcSal, " "
                Print #ArcSal, Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(83) & "Página: " & Format(vPage, "@@@@")
                Print #ArcSal, Space(1) & "Crédito Pignoraticio" & Space(81) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & "       CONSOLIDADO DE CREDITOS AL CIERRE DE MES ENTRE DIAS DE MORA "
                Print #ArcSal, Space(25) & "       =========================================================== "
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & "   TOTAL    CLASIFICACION      SALDO CAPITAL      PORCEN.    PROVISION"
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd0, 7, 0) & "   Normal            " & ImpreFormat(vSal0, 12) & ImpreFormat(pProvis0, 10, 3) & ImpreFormat(vSal0 * pProvis0, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd9, 7, 0) & "   C.P.P.            " & ImpreFormat(vSal9, 12) & ImpreFormat(pProvis9, 10, 3) & ImpreFormat(vSal9 * pProvis9, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd31, 7, 0) & "   Deficiente        " & ImpreFormat(vSal31, 12) & ImpreFormat(pProvis31, 10, 3) & ImpreFormat(vSal31 * pProvis31, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd61, 7, 0) & "   Dudoso            " & ImpreFormat(vSal61, 12) & ImpreFormat(pProvis61, 10, 3) & ImpreFormat(vSal61 * pProvis61, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd121, 7, 0) & "   Perdida           " & ImpreFormat(vSal121, 12) & ImpreFormat(pProvis121, 10, 3) & ImpreFormat(vSal121 * pProvis121, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(10) & "TOTAL" & Space(10) & ImpreFormat(vInd0 + vInd9 + vInd31 + vInd61 + vInd121, 7, 0) & _
                    Space(21) & ImpreFormat(vSal0 + vSal9 + vSal31 + vSal61 + vSal121, 12) & _
                    Space(14) & ImpreFormat(Round((vSal0 * pProvis0) + (vSal9 * pProvis9) + (vSal31 * pProvis31) + (vSal61 * pProvis61) + (vSal121 * pProvis121), 2), 10, 2)
                Print #ArcSal, " "
                ImpreEnd
                End If
            End If
        End With
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Calcula la estadística mensual
Public Function ImprimeEstadMes(ByVal poList As ListBox, poProgress As ProgressBar, _
ByVal psBoveda As String) As String
On Error GoTo ControlError
Dim pCadena As String
'Creación de ARREGLO
Dim Arreglo(24, 4) As Double
Dim vCampo As String
Dim x As Integer
Dim Ag As Integer
Dim pConexPrend As New ADODB.Connection
Dim lnBarraProg As Integer
Dim sNomAge As String

'pcadena = ""

For Ag = 1 To poList.ListCount
    If poList.Selected(Ag - 1) = True Then
        If Right(Trim(gsCodAge), 2) = Mid(poList.List(Ag - 1), 1, 2) Then
            sNomAge = gsNomAge
            Set pConexPrend = dbCmact
        Else
            If AbreConeccion(Mid(poList.List(Ag - 1), 1, 2) & "XXXXXXXXXX") Then
                sNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                Set pConexPrend = dbCmactN
            End If
        End If
                
        poProgress.Visible = True
        poProgress.Min = 0
        poProgress.Max = 100
        lnBarraProg = 0
        poProgress.Value = lnBarraProg
        
        Arreglo(1, 0) = MesNroOpe("PresNorm", 7, pConexPrend, psBoveda)
        Arreglo(2, 0) = MesNroOpe("PresDife", 7, pConexPrend, psBoveda)
        Arreglo(22, 0) = MesNroOpe("PresVencNorm", 7, pConexPrend, psBoveda)
        Arreglo(18, 0) = MesNroOpe("PresVencReno", 7, pConexPrend, psBoveda)
        Arreglo(3, 0) = MesNroOpe("PresParaRemaNorm", 7, pConexPrend, psBoveda)
        Arreglo(19, 0) = MesNroOpe("PresParaRemaReno", 7, pConexPrend, psBoveda)
        Arreglo(4, 0) = MesNroOpe("PresReno", 7, pConexPrend, psBoveda)
        Arreglo(5, 0) = Arreglo(1, 0) + Arreglo(2, 0) + Arreglo(3, 0) + Arreglo(4, 0) + Arreglo(18, 0) + Arreglo(19, 0) + Arreglo(22, 0)
        
        lnBarraProg = lnBarraProg + 10
        poProgress.Value = lnBarraProg
        
        'Suma de Lotes 14
        Arreglo(6, 0) = MesNroOpe("PresNorm", 14, pConexPrend, psBoveda)
        Arreglo(7, 0) = MesNroOpe("PresDife", 14, pConexPrend, psBoveda)
        Arreglo(23, 0) = MesNroOpe("PresVencNorm", 14, pConexPrend, psBoveda)
        Arreglo(8, 0) = MesNroOpe("PresVencReno", 14, pConexPrend, psBoveda)
        Arreglo(9, 0) = MesNroOpe("PresParaRemaNorm", 14, pConexPrend, psBoveda)
        Arreglo(10, 0) = MesNroOpe("PresParaRemaReno", 14, pConexPrend, psBoveda)
        Arreglo(11, 0) = MesNroOpe("PresReno", 14, pConexPrend, psBoveda)
        Arreglo(12, 0) = Arreglo(6, 0) + Arreglo(7, 0) + Arreglo(8, 0) + Arreglo(9, 0) + Arreglo(10, 0) + Arreglo(11, 0) + Arreglo(23, 0)
        
        lnBarraProg = lnBarraProg + 10
        poProgress.Value = lnBarraProg
        
        'Suma de Lotes 30
        Arreglo(13, 0) = MesNroOpe("PresNorm", 30, pConexPrend, psBoveda)
        Arreglo(14, 0) = MesNroOpe("PresDife", 30, pConexPrend, psBoveda)
        Arreglo(24, 0) = MesNroOpe("PresVencNorm", 30, pConexPrend, psBoveda)
        Arreglo(20, 0) = MesNroOpe("PresVencReno", 30, pConexPrend, psBoveda)
        Arreglo(15, 0) = MesNroOpe("PresParaRemaNorm", 30, pConexPrend, psBoveda)
        Arreglo(21, 0) = MesNroOpe("PresParaRemaReno", 30, pConexPrend, psBoveda)
        Arreglo(16, 0) = MesNroOpe("PresReno", 30, pConexPrend, psBoveda)
        Arreglo(17, 0) = Arreglo(13, 0) + Arreglo(14, 0) + Arreglo(15, 0) + Arreglo(16, 0) + Arreglo(20, 0) + Arreglo(21, 0) + Arreglo(24, 0)
        'Suma de Lotes TOTAL
        
        lnBarraProg = lnBarraProg + 10
        poProgress.Value = lnBarraProg
        
        Arreglo(0, 0) = Arreglo(5, 0) + Arreglo(12, 0) + Arreglo(17, 0)
        
        vCampo = "nSaldoCap"
        For x = 1 To 4
            If x = 2 Then vCampo = "nPrestamo"
            If x = 3 Then vCampo = "nValTasac"
            If x = 4 Then vCampo = "nOroNeto"
            'Suma plazo 7
            Arreglo(1, x) = MesSumOpe("PresNorm", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(2, x) = MesSumOpe("PresDife", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(22, x) = MesSumOpe("PresVencNorm", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(18, x) = MesSumOpe("PresVencReno", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(3, x) = MesSumOpe("PresParaRemaNorm", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(19, x) = MesSumOpe("PresParaRemaReno", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(4, x) = MesSumOpe("PresReno", 7, vCampo, pConexPrend, psBoveda)
            Arreglo(5, x) = Arreglo(1, x) + Arreglo(2, x) + Arreglo(3, x) + Arreglo(4, x) + Arreglo(18, x) + Arreglo(19, x) + Arreglo(22, x)
            'Suma plazo 14
            Arreglo(6, x) = MesSumOpe("PresNorm", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(7, x) = MesSumOpe("PresDife", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(23, x) = MesSumOpe("PresVencNorm", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(8, x) = MesSumOpe("PresVencReno", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(9, x) = MesSumOpe("PresParaRemaNorm", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(10, x) = MesSumOpe("PresParaRemaReno", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(11, x) = MesSumOpe("PresReno", 14, vCampo, pConexPrend, psBoveda)
            Arreglo(12, x) = Arreglo(6, x) + Arreglo(7, x) + Arreglo(8, x) + Arreglo(9, x) + Arreglo(10, x) + Arreglo(11, x) + Arreglo(23, x)
            'Suma plazo 30
            Arreglo(13, x) = MesSumOpe("PresNorm", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(14, x) = MesSumOpe("PresDife", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(24, x) = MesSumOpe("PresVencNorm", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(20, x) = MesSumOpe("PresVencReno", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(15, x) = MesSumOpe("PresParaRemaNorm", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(21, x) = MesSumOpe("PresParaRemaReno", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(16, x) = MesSumOpe("PresReno", 30, vCampo, pConexPrend, psBoveda)
            Arreglo(17, x) = Arreglo(13, x) + Arreglo(14, x) + Arreglo(15, x) + Arreglo(16, x) + Arreglo(20, x) + Arreglo(21, x) + Arreglo(24, x)
            'Suma  TOTAL de plazos
            Arreglo(0, x) = Arreglo(5, x) + Arreglo(12, x) + Arreglo(17, x)
            
            lnBarraProg = lnBarraProg + 10
            poProgress.Value = lnBarraProg
            
        Next x
        
        lnBarraProg = lnBarraProg + 10
        poProgress.Value = lnBarraProg
        
        'Carga fecha y hora de grabación
        gdHoraGrab = Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
        pCadena = pCadena & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & gsNomCmac & Space(62) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & gPrnSaltoLinea
        pCadena = pCadena & ImpreFormat(sNomAge, 45, 2) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & "Bovedas " & psBoveda & gPrnSaltoLinea
        pCadena = pCadena & Space(50) & "   ESTADISTICA   MENSUAL   " & gPrnSaltoLinea
        pCadena = pCadena & Space(50) & "(DISTRIBUCION DE CONTRATOS)" & gPrnSaltoLinea
        'pcadena = pcadena & Space(49) & String(21, "=") & gPrnSaltoLinea
        pCadena = pCadena & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & Space(30) & "LOTES        POR AMORTIZAR         PRESTAMO         V.TASACION       ORO NETO" & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  PLAZO :     7 " & gPrnSaltoLinea
        pCadena = pCadena & "  Normales                " & ImpreFormat(Arreglo(1, 0), 10, 0) & ImpreFormat(Arreglo(1, 1), 15, , True) & ImpreFormat(Arreglo(1, 2), 15, , True) & ImpreFormat(Arreglo(1, 3), 15, , True) & ImpreFormat(Arreglo(1, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Diferidos               " & ImpreFormat(Arreglo(2, 0), 10, 0) & ImpreFormat(Arreglo(2, 1), 15, , True) & ImpreFormat(Arreglo(2, 2), 15, , True) & ImpreFormat(Arreglo(2, 3), 15, , True) & ImpreFormat(Arreglo(2, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Normales       " & ImpreFormat(Arreglo(22, 0), 10, 0) & ImpreFormat(Arreglo(22, 1), 15, , True) & ImpreFormat(Arreglo(22, 2), 15, , True) & ImpreFormat(Arreglo(22, 3), 15, , True) & ImpreFormat(Arreglo(22, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Renovados      " & ImpreFormat(Arreglo(18, 0), 10, 0) & ImpreFormat(Arreglo(18, 1), 15, , True) & ImpreFormat(Arreglo(18, 2), 15, , True) & ImpreFormat(Arreglo(18, 3), 15, , True) & ImpreFormat(Arreglo(18, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Normales    " & ImpreFormat(Arreglo(3, 0), 10, 0) & ImpreFormat(Arreglo(3, 1), 15, , True) & ImpreFormat(Arreglo(3, 2), 15, , True) & ImpreFormat(Arreglo(3, 3), 15, , True) & ImpreFormat(Arreglo(3, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Renovados   " & ImpreFormat(Arreglo(19, 0), 10, 0) & ImpreFormat(Arreglo(19, 1), 15, , True) & ImpreFormat(Arreglo(19, 2), 15, , True) & ImpreFormat(Arreglo(19, 3), 15, , True) & ImpreFormat(Arreglo(19, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Renovados               " & ImpreFormat(Arreglo(4, 0), 10, 0) & ImpreFormat(Arreglo(4, 1), 15, , True) & ImpreFormat(Arreglo(4, 2), 15, , True) & ImpreFormat(Arreglo(4, 3), 15, , True) & ImpreFormat(Arreglo(4, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  Total                   " & ImpreFormat(Arreglo(5, 0), 10, 0) & ImpreFormat(Arreglo(5, 1), 15, , True) & ImpreFormat(Arreglo(5, 2), 15, , True) & ImpreFormat(Arreglo(5, 3), 15, , True) & ImpreFormat(Arreglo(5, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  PLAZO :    14 " & gPrnSaltoLinea
        pCadena = pCadena & "  Normales                " & ImpreFormat(Arreglo(6, 0), 10, 0) & ImpreFormat(Arreglo(6, 1), 15, , True) & ImpreFormat(Arreglo(6, 2), 15, , True) & ImpreFormat(Arreglo(6, 3), 15, , True) & ImpreFormat(Arreglo(6, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Diferidos               " & ImpreFormat(Arreglo(7, 0), 10, 0) & ImpreFormat(Arreglo(7, 1), 15, , True) & ImpreFormat(Arreglo(7, 2), 15, , True) & ImpreFormat(Arreglo(7, 3), 15, , True) & ImpreFormat(Arreglo(7, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Normales       " & ImpreFormat(Arreglo(23, 0), 10, 0) & ImpreFormat(Arreglo(23, 1), 15, , True) & ImpreFormat(Arreglo(23, 2), 15, , True) & ImpreFormat(Arreglo(23, 3), 15, , True) & ImpreFormat(Arreglo(23, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Renovados      " & ImpreFormat(Arreglo(8, 0), 10, 0) & ImpreFormat(Arreglo(8, 1), 15, , True) & ImpreFormat(Arreglo(8, 2), 15, , True) & ImpreFormat(Arreglo(8, 3), 15, , True) & ImpreFormat(Arreglo(8, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Normales    " & ImpreFormat(Arreglo(9, 0), 10, 0) & ImpreFormat(Arreglo(9, 1), 15, , True) & ImpreFormat(Arreglo(9, 2), 15, , True) & ImpreFormat(Arreglo(9, 3), 15, , True) & ImpreFormat(Arreglo(9, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Renovados   " & ImpreFormat(Arreglo(10, 0), 10, 0) & ImpreFormat(Arreglo(10, 1), 15, , True) & ImpreFormat(Arreglo(10, 2), 15, , True) & ImpreFormat(Arreglo(10, 3), 15, , True) & ImpreFormat(Arreglo(10, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Renovados               " & ImpreFormat(Arreglo(11, 0), 10, 0) & ImpreFormat(Arreglo(11, 1), 15, , True) & ImpreFormat(Arreglo(11, 2), 15, , True) & ImpreFormat(Arreglo(11, 3), 15, , True) & ImpreFormat(Arreglo(11, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  Total                   " & ImpreFormat(Arreglo(12, 0), 10, 0) & ImpreFormat(Arreglo(12, 1), 15, , True) & ImpreFormat(Arreglo(12, 2), 15, , True) & ImpreFormat(Arreglo(12, 3), 15, , True) & ImpreFormat(Arreglo(12, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  PLAZO :    30 " & gPrnSaltoLinea
        pCadena = pCadena & "  Normales                " & ImpreFormat(Arreglo(13, 0), 10, 0) & ImpreFormat(Arreglo(13, 1), 15, , True) & ImpreFormat(Arreglo(13, 2), 15, , True) & ImpreFormat(Arreglo(13, 3), 15, , True) & ImpreFormat(Arreglo(13, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Diferidos               " & ImpreFormat(Arreglo(14, 0), 10, 0) & ImpreFormat(Arreglo(14, 1), 15, , True) & ImpreFormat(Arreglo(14, 2), 15, , True) & ImpreFormat(Arreglo(14, 3), 15, , True) & ImpreFormat(Arreglo(14, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Normales       " & ImpreFormat(Arreglo(24, 0), 10, 0) & ImpreFormat(Arreglo(24, 1), 15, , True) & ImpreFormat(Arreglo(24, 2), 15, , True) & ImpreFormat(Arreglo(24, 3), 15, , True) & ImpreFormat(Arreglo(24, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Vencidos Renovados      " & ImpreFormat(Arreglo(20, 0), 10, 0) & ImpreFormat(Arreglo(20, 1), 15, , True) & ImpreFormat(Arreglo(20, 2), 15, , True) & ImpreFormat(Arreglo(20, 3), 15, , True) & ImpreFormat(Arreglo(20, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Normales    " & ImpreFormat(Arreglo(15, 0), 10, 0) & ImpreFormat(Arreglo(15, 1), 15, , True) & ImpreFormat(Arreglo(15, 2), 15, , True) & ImpreFormat(Arreglo(15, 3), 15, , True) & ImpreFormat(Arreglo(15, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Para Remate Renovados   " & ImpreFormat(Arreglo(21, 0), 10, 0) & ImpreFormat(Arreglo(21, 1), 15, , True) & ImpreFormat(Arreglo(21, 2), 15, , True) & ImpreFormat(Arreglo(21, 3), 15, , True) & ImpreFormat(Arreglo(21, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & "  Renovados               " & ImpreFormat(Arreglo(16, 0), 10, 0) & ImpreFormat(Arreglo(16, 1), 15, , True) & ImpreFormat(Arreglo(16, 2), 15, , True) & ImpreFormat(Arreglo(16, 3), 15, , True) & ImpreFormat(Arreglo(16, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "-") & gPrnSaltoLinea
        pCadena = pCadena & "  Total                   " & ImpreFormat(Arreglo(17, 0), 10, 0) & ImpreFormat(Arreglo(17, 1), 15, , True) & ImpreFormat(Arreglo(17, 2), 15, , True) & ImpreFormat(Arreglo(17, 3), 15, , True) & ImpreFormat(Arreglo(17, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "=") & gPrnSaltoLinea
        pCadena = pCadena & "  TOTAL                   " & ImpreFormat(Arreglo(0, 0), 10, 0) & ImpreFormat(Arreglo(0, 1), 15, , True) & ImpreFormat(Arreglo(0, 2), 15, , True) & ImpreFormat(Arreglo(0, 3), 15, , True) & ImpreFormat(Arreglo(0, 4), 15, , True) & gPrnSaltoLinea
        pCadena = pCadena & Space(2) & String(110, "=") & gPrnSaltoLinea
        'rtfImp.Text = pCadena
        'frmPrevio.Previo rtfImp, " Listado de Estadística Mensual ", True, pHojaFiMax

        poProgress.Visible = False
    End If
Next Ag
Set pConexPrend = Nothing

ImprimeEstadMes = pCadena
Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function

'Permite obtener el número de operaciones
'de la tabla de CredPrenda de acuerdo a la condición deseada
Private Function MesNroOpe(ByVal pCondicion As String, ByVal pPlazo As Integer, pConexPrend As ADODB.Connection, ByVal pBoveda As String) As Double
Dim RegProc As New ADODB.Recordset
Dim tmpSql As String
tmpSql = "SELECT count(cCodCta) AS Cuenta FROM CredPrenda "
Select Case pCondicion
    Case "PresDife"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '2' "
    Case "PresNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '1' AND nnumrenov = 0 "
    Case "PresVencNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '4' AND nnumrenov = 0 "
    Case "PresParaRemaNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '6' AND nnumrenov = 0 "
    Case "PresReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '7' AND nnumrenov > 0 "
    Case "PresVencReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '4' AND nnumrenov > 0 "
    Case "PresParaRemaReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '6' AND nnumrenov > 0 "
    Case "TODOS"

    Case Else
        MsgBox "Opción no reconocida", vbInformation, " ! Aviso ! "
        Exit Function
End Select
If Len(Trim(pBoveda)) > 0 Then
    tmpSql = tmpSql & " AND cAgeBoveda in " & pBoveda & " "
End If
RegProc.Open tmpSql, pConexPrend, adOpenForwardOnly, adLockReadOnly, adCmdText
If (RegProc.BOF Or RegProc.EOF) Then
    MesNroOpe = Format(0, "#0")
Else
    MesNroOpe = IIf(IsNull(RegProc!Cuenta) = True, 0, Format(RegProc!Cuenta, "#0"))
End If
RegProc.Close
Set RegProc = Nothing
End Function

'Permite obtener la suma de las operaciones
'de la tabla de CredPrenda de acuerdo a la condición deseada
Private Function MesSumOpe(ByVal pCondicion As String, ByVal pPlazo As Integer, ByVal pCampo As String, pConexPrend As ADODB.Connection, ByVal pBoveda As String) As Double
Dim RegProc As New ADODB.Recordset
Dim tmpSql As String
tmpSql = "SELECT Sum(" & pCampo & ") AS Suma FROM CredPrenda "
Select Case pCondicion
    Case "PresDife"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '2' "
    Case "PresNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '1' AND nnumrenov = 0 "
    Case "PresVencNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '4' AND nnumrenov = 0 "
    Case "PresParaRemaNorm"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '6' AND nnumrenov = 0 "
    Case "PresReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '7' AND nnumrenov > 0 "
    Case "PresVencReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '4' AND nnumrenov > 0 "
    Case "PresParaRemaReno"
        tmpSql = tmpSql & " WHERE nPlazo =" & pPlazo & " AND cestado = '6' AND nnumrenov > 0 "
    Case "TODOS"

    Case Else
        MsgBox "Opción no reconocida", vbInformation, " ! Aviso ! "
        Exit Function
End Select
If Len(Trim(pBoveda)) > 0 Then
    tmpSql = tmpSql & " AND cAgeBoveda in " & pBoveda & " "
End If
RegProc.Open tmpSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
If (RegProc.BOF Or RegProc.EOF) Then
    MesSumOpe = Format(0, "#0.00")
Else
    MesSumOpe = IIf(IsNull(RegProc!suma) = True, 0, Format(RegProc!suma, "#0.00"))
End If
RegProc.Close
Set RegProc = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

'Validación de txtPagIni
Private Sub txtPagIni_GotFocus()
fEnfoque txtPagIni
End Sub
Private Sub txtPagIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdImprimir.SetFocus
    'If Val(txtPagIni.Text) >= 1 And Val(txtPagIni.Text) <= vPage Then
    '    txtPagFin.Text = txtPagIni.Text
    '    txtPagFin.SetFocus
    'End If
Else
    KeyAscii = intfNumEnt(KeyAscii)
End If
End Sub
Private Sub txtPagIni_Validate(Cancel As Boolean)
'If Not (Val(txtPagIni.Text) >= 1 And Val(txtPagIni.Text) <= vPage) Then
'    Cancel = True
'End If
End Sub



'******* Modificacion
'Imprime Listado para Intereses Devengados (para Auditoria)
Private Sub ImprimeProvis2(ByVal CondicionImpresion As String, pConexPrend As ADODB.Connection)   ' Para Lourdes
' Modificado para Imprimir con Calificacion Consolidada
Dim vAtraso As Integer
Dim vPorcen As Double, vProvis  As Currency
Dim vFecVen As String
Dim vIntDev As Currency
Dim vNombre As String * 35
Dim vIndice As Double  'contador de Item
Dim vLineas As Integer
Dim vPage As Integer
Dim vTotSald As Currency, vTotVaTa As Currency
Dim vTotIntDev As Currency
Dim vTotOrNe As Currency, vTotProvis As Currency
Dim vCabecera As String
Dim vCont As Double
Dim vInd0 As Double, vSal0 As Currency
Dim vInd9 As Double, vSal9 As Currency
Dim vInd31 As Double, vSal31 As Currency
Dim vInd61 As Double, vSal61 As Currency
Dim vInd121 As Double, vSal121 As Currency
Dim v0 As Boolean, v9 As Boolean, v31 As Boolean, v61 As Boolean, v121 As Boolean
    MousePointer = 11
    MuestraImpresion = True
    sSql = "SELECT cp.cCodCta, cp.dfecpres, cp.nplazo, cp.nprestamo, cp.nsaldocap, " & _
        " cp.nValTasac, cp.nOroNeto, cp.dfecvenc, cp.ntasaint, CA.cCalGen, " & _
        " datediff(dd,cp.dfecvenc, '" & Format(gdFecSis, "mm/dd/yyyy") & "') AS DiasDife " & _
        " FROM CredPrenda CP LEFT JOIN CREDITOAUDI CA ON CP.CCODCTA = CA.CCODCTA " & _
        " WHERE cp.cestado IN ('1','4','7','6') " & _
        " ORDER BY CA.CCALGEN, DiasDife, cp.dfecpres, cp.cCodCta "
    vCabecera = "Provis2"
    RegCredPrend.Open sSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        Select Case CondicionImpresion
            Case "Provis"
                MsgBox " No existen Contratos ", vbInformation, " ! Aviso ! "
        End Select
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegCredPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegCredPrend.Close
                Set RegCredPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        ElseIf optImpresion(1).Value = True Then
            ImpreBegin True, pHojaFiMax
            vRTFImp = ""
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
                Print #ArcSal, ImpreCarEsp(vRTFImp);
                vRTFImp = ""
            End If
        Else
            If Val(txtPagIni) <= 1 Then
                Cabecera vCabecera, vPage, False
            End If
        End If
        prgList.Visible = True
        vPorcen = 0: vProvis = 0
        vTotVaTa = 0: vTotOrNe = 0: vTotProvis = 0
        v0 = True: v9 = True: v31 = True: v61 = True: v121 = True
        vInd0 = 0: vInd9 = 0: vInd31 = 0: vInd61 = 0: vInd121 = 0:
        vSal0 = 0: vSal9 = 0: vSal31 = 0: vSal61 = 0: vSal121 = 0:
        vIndice = 1
        vLineas = 7
        With RegCredPrend
            Do While Not .EOF
                If !cCalGen = "0" Then
                    If v0 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   N O R M A L                " & pCorte
                        End If
                        v0 = False
                        vPorcen = pProvis0
                    End If
                    vInd0 = vInd0 + 1: vSal0 = vSal0 + !nSaldoCap
                ElseIf !cCalGen = "1" Then
                    If v9 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   PROBLEMAS POTENCIALES " & pCorte
                        End If
                        v9 = False
                        vPorcen = pProvis9
                    End If
                    vInd9 = vInd9 + 1: vSal9 = vSal9 + !nSaldoCap
                ElseIf !cCalGen = "2" Then
                    If v31 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DEFICIENTE " & pCorte
                        End If
                        v31 = False
                        vPorcen = pProvis31
                    End If
                    vInd31 = vInd31 + 1: vSal31 = vSal31 + !nSaldoCap
                ElseIf !cCalGen = "3" Then
                    If v61 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   DUDOSO  " & pCorte
                        End If
                        v61 = False
                        vPorcen = pProvis61
                    End If
                    vInd61 = vInd61 + 1: vSal61 = vSal61 + !nSaldoCap
                ElseIf !cCalGen = "4" Then
                    If v121 Then
                        If Len(Trim(vRTFImp)) > 0 Then
                        vRTFImp = vRTFImp & pCorte & "   PERDIDA " & pCorte
                        End If
                        v121 = False
                        vPorcen = pProvis121
                    End If
                    vInd121 = vInd121 + 1: vSal121 = vSal121 + !nSaldoCap
                End If
                vAtraso = 0
                vFecVen = Format(!dFecVenc, "dd/mm/yyyy")
                vNombre = PstaNombre(ClienteNombre(!cCodCta, pConexPrend), False)
                If DateDiff("d", vFecVen, gdFecSis) > 0 Then
                    vIntDev = ((1 + pTasaInteresVencido) ^ (DateDiff("d", vFecVen, gdFecSis) / 30) - 1) * !nSaldoCap
                    vIntDev = Round(vIntDev, 2)
                End If
                vAtraso = IIf(DateDiff("d", vFecVen, gdFecSis) > 0, DateDiff("d", vFecVen, gdFecSis), 0)
                vProvis = !nSaldoCap * vPorcen
                vProvis = Round(vProvis, 2)
                If vIndice > ((Val(txtPagIni) * pItemMax) - pItemMax) Then
                If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 8, 0) & ImpreFormat(!cCodCta, 13, 1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreFormat(vNombre, 25, 1) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 4, 0) & Space(1) & Format(!dFecVenc, "dd/mm/yyyy") & _
                        ImpreFormat(vIntDev, 7) & ImpreFormat(vAtraso, 6, 0) & _
                        ImpreFormat(!nSaldoCap + vIntDev, 10) & ImpreFormat(!noroneto, 7) & _
                        ImpreFormat(vPorcen, 5, 3) & ImpreFormat(vProvis, 8) & pCorte
                    If vIndice Mod 300 = 0 Then
                        vBuffer = vBuffer & vRTFImp
                        vRTFImp = ""
                    End If
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 8, 0) & ImpreFormat(!cCodCta, 13, 1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreCarEsp(ImpreFormat(vNombre, 25, 1)) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 4, 0) & Space(1) & Format(!dFecVenc, "dd/mm/yyyy") & _
                        ImpreFormat(vIntDev, 7) & ImpreFormat(vAtraso, 6, 0) & _
                        ImpreFormat(!nSaldoCap + vIntDev, 10) & ImpreFormat(!noroneto, 7) & _
                        ImpreFormat(vPorcen, 5, 3) & ImpreFormat(vProvis, 8)
                End If
                End If
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                vTotSald = vTotSald + !nSaldoCap
                vTotVaTa = vTotVaTa + !nvaltasac
                vTotIntDev = vTotIntDev + vIntDev
                vTotOrNe = vTotOrNe + !noroneto
                vTotProvis = vTotProvis + vProvis
                
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If vPage > (Val(txtPagIni) - 1) Then
                    If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                        If Val(txtPagIni) <> vPage Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        End If
                        Cabecera vCabecera, vPage, False
                    Else
                        If Val(txtPagIni) <> vPage Then
                            If vPage Mod 5 = 0 Then
                                ImpreEnd
                                ImpreBegin True, pHojaFiMax
                            Else
                                ImpreNewPage
                            End If
                        End If
                        vRTFImp = ""
                        Cabecera vCabecera, vPage, False
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotSald = Round(vTotSald, 2)
            If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
                If Val(txtPagIni) <= vPage Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 40) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 22) & _
                    ImpreFormat(vTotSald + vTotIntDev, 16) & ImpreFormat(Round(vTotOrNe, 2), 7) & _
                    ImpreFormat(Round(vTotProvis, 2), 17) & pCorte
                vRTFImp = vRTFImp & gPrnSaltoPagina
                End If
                vPage = vPage + 1
                If Val(txtPagIni) <= vPage Then
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(83) & "Página: " & Format(vPage, "@@@@") & pCorte
                vRTFImp = vRTFImp & Space(1) & "Crédito Pignoraticio" & Space(81) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & "       CONSOLIDADO DE CREDITOS AL CIERRE DE MES ENTRE DIAS DE MORA " & pCorte
                vRTFImp = vRTFImp & Space(25) & "       =========================================================== " & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & "   TOTAL    CLASIFICACION      SALDO CAPITAL      PORCEN.    PROVISION" & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd0, 7, 0) & "   Normal            " & ImpreFormat(vSal0, 12) & ImpreFormat(pProvis0, 10, 3) & ImpreFormat(vSal0 * pProvis0, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd9, 7, 0) & "   Problemas Pot.    " & ImpreFormat(vSal9, 12) & ImpreFormat(pProvis9, 10, 3) & ImpreFormat(vSal9 * pProvis9, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd31, 7, 0) & "   Deficiente        " & ImpreFormat(vSal31, 12) & ImpreFormat(pProvis31, 10, 3) & ImpreFormat(vSal31 * pProvis31, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd61, 7, 0) & "   Dudoso            " & ImpreFormat(vSal61, 12) & ImpreFormat(pProvis61, 10, 3) & ImpreFormat(vSal61 * pProvis61, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(25) & ImpreFormat(vInd121, 7, 0) & "   Perdida           " & ImpreFormat(vSal121, 12) & ImpreFormat(pProvis121, 10, 3) & ImpreFormat(vSal121 * pProvis121, 10, 3) & pCorte
                vRTFImp = vRTFImp & pCorte
                vRTFImp = vRTFImp & Space(10) & "TOTAL" & Space(10) & ImpreFormat(vInd0 + vInd9 + vInd31 + vInd61 + vInd121, 7, 0) & _
                    Space(21) & ImpreFormat(vSal0 + vSal9 + vSal31 + vSal61 + vSal121, 12) & _
                    Space(14) & ImpreFormat(Round((vSal0 * pProvis0) + (vSal9 * pProvis9) + (vSal31 * pProvis31) + (vSal61 * pProvis61) + (vSal121 * pProvis121), 2), 10, 2) & pCorte
                vRTFImp = vRTFImp & pCorte
                End If
                vBuffer = vBuffer & vRTFImp
                vRTFImp = ""
            Else
                If Val(txtPagIni) <= vPage Then
                Print #ArcSal, " "
                Print #ArcSal, ImpreFormat("TOTAL", 6, 14) & ImpreFormat((vIndice - 1), 8, 0) & _
                    ImpreFormat(vTotSald, 40) & ImpreFormat(vTotVaTa, 10) & _
                    ImpreFormat(vTotIntDev, 22) & _
                    ImpreFormat(vTotSald + vTotIntDev, 16) & ImpreFormat(Round(vTotOrNe, 2), 7) & _
                    ImpreFormat(vTotProvis, 17)
                ImpreNewPage
                End If
                vPage = vPage + 1
                If Val(txtPagIni) <= vPage Then
                Print #ArcSal, " "
                Print #ArcSal, Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(83) & "Página: " & Format(vPage, "@@@@")
                Print #ArcSal, Space(1) & "Crédito Pignoraticio" & Space(81) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & "       CONSOLIDADO DE CREDITOS AL CIERRE DE MES ENTRE DIAS DE MORA "
                Print #ArcSal, Space(25) & "       =========================================================== "
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & "   TOTAL    CLASIFICACION      SALDO CAPITAL      PORCEN.    PROVISION"
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd0, 7, 0) & "   Normal            " & ImpreFormat(vSal0, 12) & ImpreFormat(pProvis0, 10, 3) & ImpreFormat(vSal0 * pProvis0, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd9, 7, 0) & "   C.P.P.            " & ImpreFormat(vSal9, 12) & ImpreFormat(pProvis9, 10, 3) & ImpreFormat(vSal9 * pProvis9, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd31, 7, 0) & "   Deficiente        " & ImpreFormat(vSal31, 12) & ImpreFormat(pProvis31, 10, 3) & ImpreFormat(vSal31 * pProvis31, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd61, 7, 0) & "   Dudoso            " & ImpreFormat(vSal61, 12) & ImpreFormat(pProvis61, 10, 3) & ImpreFormat(vSal61 * pProvis61, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(25) & ImpreFormat(vInd121, 7, 0) & "   Perdida           " & ImpreFormat(vSal121, 12) & ImpreFormat(pProvis121, 10, 3) & ImpreFormat(vSal121 * pProvis121, 10, 3)
                Print #ArcSal, " "
                Print #ArcSal, Space(10) & "TOTAL" & Space(10) & ImpreFormat(vInd0 + vInd9 + vInd31 + vInd61 + vInd121, 7, 0) & _
                    Space(21) & ImpreFormat(vSal0 + vSal9 + vSal31 + vSal61 + vSal121, 12) & _
                    Space(14) & ImpreFormat(Round((vSal0 * pProvis0) + (vSal9 * pProvis9) + (vSal31 * pProvis31) + (vSal61 * pProvis61) + (vSal121 * pProvis121), 2), 10, 2)
                Print #ArcSal, " "
                ImpreEnd
                End If
            End If
        End With
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

Private Function MuestraCodigoContrato(pContrato As String) As String

End Function
Private Function FormatoContratro(pContrato As String) As String
FormatoContratro = Mid(pContrato, 1, 2) & "-" & Mid(pContrato, 3, 4) & "-" & Mid(pContrato, 7, 5) & "-" & Mid(pContrato, 12, 1)
End Function
