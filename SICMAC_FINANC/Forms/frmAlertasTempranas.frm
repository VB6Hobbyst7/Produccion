VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAlertasTempranas 
   Caption         =   "Alertas Tempranas"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frmAlertasTempranas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   4200
      TabIndex        =   49
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame fraProceso 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   9735
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   8520
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   6240
         TabIndex        =   27
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5520
         TabIndex        =   28
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Límites Regulatorios, Límites Internos y Alertas Tempranas de Liquidez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Label lblActivoTotales 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   48
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label lblActivosTOSE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   47
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbl10 
         Caption         =   "10"
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
         Left            =   4440
         TabIndex        =   46
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lbl11 
         Caption         =   "11"
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
         Left            =   4440
         TabIndex        =   45
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Activos Líquidos /Activos Totales exp en MN"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   4320
         Width           =   3975
      End
      Begin VB.Label Label10 
         Caption         =   "Activos Líquidos / TOSE exp en Soles"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label lblAct6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblAct5 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   41
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblAct4 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblAct3 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblAct2 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblAct1 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblPCL 
         Alignment       =   2  'Center
         Caption         =   "ACTIVACION DE PCL"
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
         Left            =   8040
         TabIndex        =   36
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Encaje Exigible / Activos Líquidos ME"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Encaje Exigible / Activos Líquidos MN"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label lblEncajeME 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   33
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lblEncajeMN 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   32
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lbl9 
         Caption         =   "9"
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
         Left            =   4440
         TabIndex        =   31
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lbl8 
         Caption         =   "8"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblRatios 
         Caption         =   "Ratios"
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Ratio de Cobertura de Liquidez ME"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Ratio de Cobertura de Liquidez MN"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Ratio de Inversiones Líquidas MN"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Ratio de Liquidez Ajustado por Recursos Prestado ME "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Ratio de Liquidez Ajustado por Recursos Prestado MN "
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Rato de Liquidez ME"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Rato de Liquidez MN"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lbl7 
         Caption         =   "7"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "6"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lbl5 
         Caption         =   "5"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbl4 
         Caption         =   "4"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lbl3 
         Caption         =   "3"
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
         Left            =   4440
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbl2 
         Caption         =   "2"
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
         Left            =   4440
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl1 
         Caption         =   "1"
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
         Left            =   4440
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblResultados 
         Caption         =   "Resultados"
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
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNivel 
         Caption         =   "Nivel de Riesgo Asumido "
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
         Left            =   5280
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblRatioCLME 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   7
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblRatioCLMN 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblRatioILMN 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblRatioLAME 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRatioLAMN 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblRatioLME 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblRatioLMN 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAlertasTempranas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JIPR20200824
Private Sub cmActualizar_Click()
Dim psArchivoALeer As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bexiste, bencontrado As Boolean
Dim pdFecha As Date
Dim fs As New Scripting.FileSystemObject
Dim rsDatos As ADODB.Recordset
Dim rsNiveles As ADODB.Recordset
Dim oNiveles As frmAlertasTempranas
Set oNiveles = New frmAlertasTempranas

 If Not IsDate(txtFecha) Then
    MsgBox "Formato de Fecha no válido...!!", vbInformation, "SICMACM - Aviso"
    txtFecha.SetFocus
    Exit Sub
 ElseIf CDate(txtFecha) > gdFecSis Then
    MsgBox "Por favor ingrese la fecha correspondiente.", vbInformation, "SICMACM - Aviso"
    txtFecha.SetFocus
    Exit Sub
 End If
 If CDate(txtFecha) < 0 Then
            MsgBox "Debe Ingresar una fecha para Saldos a Favor", vbInformation, "SICMACM - Aviso"
            txtFecha.SetFocus
            Exit Sub
 End If

pdFecha = txtFecha.Text
psArchivoALeer = App.path & "\SPOOLER\" & "Anx15A_New_" & Format(pdFecha, "yyyymmdd") & ".xlsx"
bexiste = fs.FileExists(psArchivoALeer)

If bexiste = False Then
    psArchivoALeer = App.path & "\SPOOLER\" & "Anx15A_New_" & Format(pdFecha, "yyyymmdd") & ".xlsx"
    bexiste = fs.FileExists(psArchivoALeer)
    If bexiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    End If
Else
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bencontrado = True
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase("validación") Then
            bencontrado = True
            xlHoja.Activate
            Exit For
'        End If
        Else
        bencontrado = False
        End If
    Next
    
    If bencontrado = True Then
    
lbl1.Caption = Format(xlHoja.Cells(233, 6), "#,##0.00")
lbl2.Caption = Format(xlHoja.Cells(234, 6), "#,##0.00")
lbl3.Caption = Format(xlHoja.Cells(235, 6), "#,##0.00")
lbl4.Caption = Format(xlHoja.Cells(236, 6), "#,##0.00")
lbl5.Caption = Format(xlHoja.Cells(237, 6), "#,##0.00")
lbl6.Caption = Format(xlHoja.Cells(238, 6), "#,##0.00")
lbl7.Caption = Format(xlHoja.Cells(239, 6), "#,##0.00")
lbl8.Caption = Format(xlHoja.Cells(240, 6), "#,##0.00")
lbl9.Caption = Format(xlHoja.Cells(241, 6), "#,##0.00")
lbl10.Caption = Format(xlHoja.Cells(242, 6), "#,##0.00")
lbl11.Caption = Format(xlHoja.Cells(243, 6), "#,##0.00")

    Set rsNiveles = oNiveles.SeleccionaNiveles
                
                    Do While Not rsNiveles.EOF
                    
                    psBajo = rsNiveles!cBajo
                    psModerado = rsNiveles!cModerado
                    psAlto = rsNiveles!cAlto
                    psExtremo = rsNiveles!cExtremo
                    psActivacionPCL = rsNiveles!cActPCL
                    
                    
                    Select Case rsNiveles!nIdNivelRgo
                    
                           Case 1
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(233, 6), "#,##0.00"))
                                If nResultAUX > CCur(psBajo) Then
                                    lblRatioLMN.BackColor = &HC000&
                                    lblRatioLMN.Caption = "BAJO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblRatioLMN.BackColor = &HFFFF&
                                    lblRatioLMN.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                    lblRatioLMN.BackColor = &H80FF&
                                    lblRatioLMN.Caption = "ALTO"
                                'End If

                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioLMN.BackColor = &HFF&
                                    lblRatioLMN.Caption = "EXTREMO"
'                                End If
                                Else
                                    lblRatioLMN.Caption = ""
                                    lblRatioLMN.BackColor = &H8000000E
                                End If
                                
                                If nResultAUX >= CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
                                lblAct1.Caption = "SE ACTIVA PCL"
                                                         
                                 Else
                                    lblAct1.Caption = ""
                                    lblAct1.BackColor = &H8000000E
                                End If
                                

                           Case 2
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(234, 6), "#,##0.00"))
                                If nResultAUX > CCur(psBajo) Then
                                 lblRatioLME.BackColor = &HC000&
                                 lblRatioLME.Caption = "BAJO"
                               ' End If
                                ElseIf nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblRatioLME.BackColor = &HFFFF&
                                    lblRatioLME.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                    lblRatioLME.BackColor = &H80FF&
                                    lblRatioLME.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioLME.BackColor = &HFF&
                                    lblRatioLME.Caption = "EXTREMO"
                              
                                Else
                                    lblRatioLME.Caption = ""
                                    lblRatioLME.BackColor = &H8000000E
                                End If
                                
                                If nResultAUX >= CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
                                    lblAct2.Caption = "SE ACTIVA PCL"
                              
                                 Else
                                    lblAct2.Caption = ""
                                    lblAct2.BackColor = &H8000000E
                                End If
                                
                           Case 3
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(235, 6), "#,##0.00"))
                                If nResultAUX > CCur(psBajo) Then
                                    lblRatioLAMN.BackColor = &HC000&
                                    lblRatioLAMN.Caption = "BAJO"
                                'End If
                                 ElseIf nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblRatioLAMN.BackColor = &HFFFF&
                                    lblRatioLAMN.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                    lblRatioLAMN.BackColor = &H80FF&
                                    lblRatioLAMN.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioLAMN.BackColor = &HFF&
                                    lblRatioLAMN.Caption = "EXTREMO"
                              
                                Else
                                    lblRatioLAMN.Caption = ""
                                    lblRatioLAMN.BackColor = &HFFF&
                                End If
                                
                                If nResultAUX > CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
                                    lblAct3.Caption = "SE ACTIVA PCL"
                             
                                 Else
                                    lblAct3.Caption = ""
                                    lblAct3.BackColor = &H8000000E
                                End If
                                
                           Case 4
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(236, 6), "#,##0.00"))
                                If nResultAUX > CCur(psBajo) Then
                                    lblRatioLAME.BackColor = &HC000&
                                    lblRatioLAME.Caption = "BAJO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblRatioLAME.BackColor = &HFFFF&
                                    lblRatioLAME.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                    lblRatioLAME.BackColor = &H80FF&
                                    lblRatioLAME.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioLAME.BackColor = &HFF&
                                    lblRatioLAME.Caption = "EXTREMO"
                             
                                Else
                                    lblRatioLAME.Caption = ""
                                    lblRatioLAME.BackColor = &H8000000E
                                End If
                                
                                If nResultAUX > CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
                                lblAct4.Caption = "SE ACTIVA PCL"
                              
                                Else
                                    lblAct4.Caption = ""
                                    lblAct4.BackColor = &H8000000E
                                End If
                                
                           Case 5
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(237, 6), "#,##0.00"))
                                If nResultAUX > CCur(psBajo) Then
                                    lblRatioILMN.BackColor = &HC000&
                                    lblRatioILMN.Caption = "BAJO"
                                'End If
                                 ElseIf nResultAUX > CCur(Left(psModerado, 1)) And nResultAUX <= CCur(Right(psModerado, 1)) Then
                                    lblRatioILMN.BackColor = &HFFFF&
                                    lblRatioILMN.Caption = "MODERADO"
                                'End If
                                 ElseIf nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 1)) Then
                                    lblRatioILMN.BackColor = &H80FF&
                                    lblRatioILMN.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioILMN.BackColor = &HFF&
                                    lblRatioILMN.Caption = "EXTREMO"
                                
                                Else
                                    lblRatioILMN.Caption = ""
                                    lblRatioILMN.BackColor = &H8000000E
                                End If
                                
                           Case 6
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(238, 6), "#,##0.00"))
                                If nResultAUX > CCur(Left(psBajo, 3)) And nResultAUX <= CCur(Right(psBajo, 3)) Then
                                    lblRatioCLMN.BackColor = &HC000&
                                    lblRatioCLMN.Caption = "BAJO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psModerado, 3)) And nResultAUX <= CCur(Right(psModerado, 3)) Then
                                     lblRatioCLMN.BackColor = &HFFFF&
                                     lblRatioCLMN.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 3)) And nResultAUX <= CCur(Right(psAlto, 3)) Then
                                    lblRatioCLMN.BackColor = &H80FF&
                                    lblRatioCLMN.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioCLMN.BackColor = &HFF&
                                    lblRatioCLMN.Caption = "EXTREMO"
                            
                                Else
                                    lblRatioCLMN.Caption = ""
                                    lblRatioCLMN.BackColor = &H8000000E
                                End If
                                
                                If nResultAUX > CCur(Left(psActivacionPCL, 3)) And nResultAUX <= CCur(Right(psActivacionPCL, 3)) Then
                                    lblAct5.Caption = "SE ACTIVA PCL"
                            
                                Else
                                    lblAct5.Caption = ""
                                    lblAct5.BackColor = &H8000000E
                                End If
                                
                           Case 7
                           
                                nResultAUX = CCur(Format(xlHoja.Cells(239, 6), "#,##0.00"))
                                If nResultAUX > CCur(Left(psBajo, 3)) And nResultAUX <= CCur(Right(psBajo, 3)) Then
                                    lblRatioCLME.BackColor = &HC000&
                                    lblRatioCLME.Caption = "BAJO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psModerado, 3)) And nResultAUX <= CCur(Right(psModerado, 3)) Then
                                    lblRatioCLME.BackColor = &HFFFF&
                                    lblRatioCLME.Caption = "MODERADO"
                                'End If
                                ElseIf nResultAUX > CCur(Left(psAlto, 3)) And nResultAUX <= CCur(Right(psAlto, 3)) Then
                                    lblRatioCLME.BackColor = &H80FF&
                                    lblRatioCLME.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblRatioCLME.BackColor = &HFF&
                                    lblRatioCLME.Caption = "EXTREMO"
                            
                                Else
                                    lblRatioCLME.Caption = ""
                                    lblRatioCLME.BackColor = &H8000000E
                                End If
                                
                                If nResultAUX > CCur(Left(psActivacionPCL, 3)) And nResultAUX <= CCur(Right(psActivacionPCL, 3)) Then
                                    lblAct6.Caption = "SE ACTIVA PCL"
                             
                                Else
                                    lblAct6.Caption = ""
                                    lblAct6.BackColor = &H8000000E
                                End If
                                
                          Case 8
                          
                                nResultAUX = CCur(Format(xlHoja.Cells(240, 6), "#,##0.00"))
                                If nResultAUX > CCur(Left(psBajo, 2)) And nResultAUX < CCur(Right(psBajo, 2)) Then
                                    lblEncajeMN.BackColor = &HC000&
                                    lblEncajeMN.Caption = "BAJO"
                                'End If
                                
                                ElseIf (nResultAUX >= CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Mid(psModerado, 3, 2))) Or (nResultAUX >= CCur(Mid(psModerado, 6, 2)) And nResultAUX <= CCur(Mid(psModerado, 8, 2))) Then
                                    lblEncajeMN.BackColor = &HFFFF&
                                    lblEncajeMN.Caption = "MODERADO"
                               ' End If
                                
                               ElseIf (nResultAUX >= CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Mid(psAlto, 3, 2))) Or (nResultAUX >= CCur(Mid(psAlto, 6, 2)) And nResultAUX <= CCur(Mid(psAlto, 8, 2))) Then
                                    lblEncajeMN.BackColor = &H80FF&
                                    lblEncajeMN.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX < CCur(Left(psExtremo, 2)) Or nResultAUX > CCur(Right(psExtremo, 2)) Then
                                    lblEncajeMN.BackColor = &HFF&
                                    lblEncajeMN.Caption = "EXTREMO"
                             
                                Else
                                    lblEncajeMN.Caption = ""
                                    lblEncajeMN.BackColor = &H8000000E
                                End If
                            
                          Case 9
                                nResultAUX = CCur(Format(xlHoja.Cells(241, 6), "#,##0.00"))
                                If nResultAUX > CCur(Left(psBajo, 2)) And nResultAUX < CCur(Right(psBajo, 2)) Then
                                    lblEncajeME.BackColor = &HC000&
                                    lblEncajeME.Caption = "BAJO"
                                'End If
                                ElseIf (nResultAUX >= CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Mid(psModerado, 3, 2))) Or (nResultAUX >= CCur(Mid(psModerado, 6, 2)) And nResultAUX <= CCur(Mid(psModerado, 8, 2))) Then
                                    lblEncajeME.BackColor = &HFFFF&
                                    lblEncajeME.Caption = "MODERADO"
                                'End If
                                ElseIf (nResultAUX >= CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Mid(psAlto, 3, 2))) Or (nResultAUX >= CCur(Mid(psAlto, 6, 2)) And nResultAUX <= CCur(Mid(psAlto, 8, 2))) Then
                                    lblEncajeME.BackColor = &H80FF&
                                    lblEncajeME.Caption = "ALTO"
                                'End If
                                ElseIf nResultAUX < CCur(Left(psExtremo, 2)) Or nResultAUX > CCur(Right(psExtremo, 2)) Then
                                  lblEncajeME.BackColor = &HFF&
                                  lblEncajeME.Caption = "EXTREMO"
                             
                                Else
                                    lblEncajeME.Caption = ""
                                    lblEncajeME.BackColor = &H8000000E
                                End If
                            
                         Case 12
                                 nResultAUX = CCur(Format(xlHoja.Cells(242, 6), "#,##0.00"))
                                 If nResultAUX > CCur(psBajo) Then
                                     lblActivosTOSE.BackColor = &HC000&
                                     lblActivosTOSE.Caption = "BAJO"
                                 'End If
                                 ElseIf nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblActivosTOSE.BackColor = &HFFFF&
                                    lblActivosTOSE.Caption = "MODERADO"
                                 'End If
                                 ElseIf nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                    lblActivosTOSE.BackColor = &H80FF&
                                    lblActivosTOSE.Caption = "ALTO"
                                 'End If
                                 ElseIf nResultAUX <= CCur(psExtremo) Then
                                    lblActivosTOSE.BackColor = &HFF&
                                    lblActivosTOSE.Caption = "EXTREMO"
                             
                                Else
                                    lblActivosTOSE.Caption = ""
                                    lblActivosTOSE.BackColor = &H8000000E
                                End If
                                 
                            Case 13
                                 nResultAUX = CCur(Format(xlHoja.Cells(243, 6), "#,##0.00"))
                                 If nResultAUX > CCur(psBajo) Then
                                     lblActivoTotales.BackColor = &HC000&
                                     lblActivoTotales.Caption = "BAJO"
                                 'End If
                                 ElseIf nResultAUX > CCur(Left(psModerado, 4)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblActivoTotales.BackColor = &HFFFF&
                                    lblActivoTotales.Caption = "MODERADO"
                                 'End If
                                 ElseIf nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 4)) Then
                                    lblActivoTotales.BackColor = &H80FF&
                                    lblActivoTotales.Caption = "ALTO"
                                 'End If
                                 ElseIf nResultAUX > CCur(Left(psExtremo, 4)) And nResultAUX < CCur(Right(psExtremo, 2)) Then
                                    lblActivoTotales.BackColor = &HFF&
                                    lblActivoTotales.Caption = "EXTREMO"
                             
                                Else
                                    lblActivoTotales.Caption = ""
                                    lblActivoTotales.BackColor = &H8000000E
                                End If

                     End Select
                     
                    rsNiveles.MoveNext
                    Loop
    Else
     MsgBox "No se encuentra la pestaña validación", vbExclamation, "Aviso!!!"
    End If
End If
End Sub
'JIPR20200824

Private Sub cmdLimpiar_Click()
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lbl10.Caption = "" 'NAGL 20190624
lbl11.Caption = "" 'JIPR20201023
lblRatioLMN.BackColor = &H8000000E
lblRatioLME.BackColor = &H8000000E
lblRatioLAMN.BackColor = &H8000000E
lblRatioLAME.BackColor = &H8000000E
lblRatioILMN.BackColor = &H8000000E
lblRatioCLMN.BackColor = &H8000000E
lblRatioCLME.BackColor = &H8000000E
lblEncajeMN.BackColor = &H8000000E
lblEncajeME.BackColor = &H8000000E
lblActivosTOSE.BackColor = &H8000000E 'NAGL 20190624
lblActivoTotales.BackColor = &H8000000E 'JIPR20200824
lblRatioLMN.Caption = ""
lblRatioLME.Caption = ""
lblRatioLAMN.Caption = ""
lblRatioLAME.Caption = ""
lblRatioILMN.Caption = ""
lblRatioCLMN.Caption = ""
lblRatioCLME.Caption = ""
lblEncajeMN.Caption = ""
lblEncajeME.Caption = ""
lblActivosTOSE.Caption = "" 'NAGL 20190624
lblActivoTotales.Caption = "" 'JIPR20200824
End Sub

Private Sub cmProcesar_Click()
Dim rsDatos As ADODB.Recordset
Dim rsNiveles As ADODB.Recordset
Dim oAlertas As frmAlertasTempranas
Set oAlertas = New frmAlertasTempranas
Dim oNiveles As frmAlertasTempranas
Set oNiveles = New frmAlertasTempranas
Dim psBajo, psModerado, psAlto, psExtremo, psActivacionPCL As String
Dim nResultAUX As Double
Dim psColor As String

 If Not IsDate(txtFecha) Then
    MsgBox "Formato de Fecha no válido...!!", vbInformation, "SICMACM - Aviso"
    txtFecha.SetFocus
    Exit Sub
 ElseIf CDate(txtFecha) > gdFecSis Then
    MsgBox "Por favor ingrese la fecha correspondiente.", vbInformation, "SICMACM - Aviso"
    txtFecha.SetFocus
    Exit Sub
 End If
 If CDate(txtFecha) < 0 Then
            MsgBox "Debe Ingresar una fecha para Saldos a Favor", vbInformation, "SICMACM - Aviso"
            txtFecha.SetFocus
            Exit Sub
 End If
       Set rsDatos = oAlertas.SeleccionaAlertasTempranA(txtFecha)
       
        If Not (rsDatos.BOF Or rsDatos.EOF) Then
        
          lbl1.Caption = Format(rsDatos!nRatioLMN, "#,##0.00")
          lbl2.Caption = Format(rsDatos!nRatioLME, "#,##0.00")
          lbl3.Caption = Format(rsDatos!nRatioLAMN, "#,##0.00")
          lbl4.Caption = Format(rsDatos!nRatioLAME, "#,##0.00")
          lbl5.Caption = Format(rsDatos!nRatioILMN, "#,##0.00")
          lbl6.Caption = Format(rsDatos!nRatioCLMN, "#,##0.00")
          lbl7.Caption = Format(rsDatos!nRatioCLME, "#,##0.00")
          lbl8.Caption = Format(rsDatos!nEncajeMN, "#,##0.00")
          lbl9.Caption = Format(rsDatos!nEncajeME, "#,##0.00")
          lbl10.Caption = Format(rsDatos!nActivosTOSE, "#,##0.00") 'NAGL Anx02_ERS006-2019
          lbl11.Caption = Format(rsDatos!nActivosTotales, "#,#0.00")
            
                Set rsNiveles = oNiveles.SeleccionaNiveles
                
                    Do While Not rsNiveles.EOF
                    
                    psBajo = rsNiveles!cBajo
                    psModerado = rsNiveles!cModerado
                    psAlto = rsNiveles!cAlto
                    psExtremo = rsNiveles!cExtremo
                    psActivacionPCL = rsNiveles!cActPCL
                    
                    
                    Select Case rsNiveles!nIdNivelRgo
                    
                           Case 1
                                nResultAUX = CCur(rsDatos!nRatioLMN) 'JIPR20201023
                                'nResultAUX = rsDatos!nRatioLMN
                                If nResultAUX > psBajo Then
                                    lblRatioLMN.BackColor = &HC000&
                                    lblRatioLMN.Caption = "BAJO"
                                End If
                                If nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20200824
                                    lblRatioLMN.BackColor = &HFFFF&
                                    lblRatioLMN.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                               ' If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20200824
                                    lblRatioLMN.BackColor = &H80FF&
                                    lblRatioLMN.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
                                'If nResultAUX <= psExtremo Then JIPR20200824
                                    lblRatioLMN.BackColor = &HFF&
                                    lblRatioLMN.Caption = "EXTREMO"
                                End If
                                If nResultAUX >= CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
                                'If nResultAUX >= Left(psActivacionPCL, 2) And nResultAUX <= Mid(psActivacionPCL, 3, 2) Then JIPR20200824
                                    'lblAct1.Caption = "ACTIVO" JIPR20200824
                                    lblAct1.Caption = "SE ACTIVA PCL"
                                End If
                                
                           Case 2
                           
                                nResultAUX = CCur(rsDatos!nRatioLME)
                                If nResultAUX > CCur(psBajo) Then
                                'nResultAUX = rsDatos!nRatioLME JIPR20200824
                                'If nResultAUX > psBajo Then
                                 lblRatioLME.BackColor = &HC000&
                                 lblRatioLME.Caption = "BAJO"
                                End If
                                If nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20200824
                                    lblRatioLME.BackColor = &HFFFF&
                                    lblRatioLME.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20200824
                                    lblRatioLME.BackColor = &H80FF&
                                    lblRatioLME.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
'                                If nResultAUX <= psExtremo Then JIPR20200824
                                    lblRatioLME.BackColor = &HFF&
                                    lblRatioLME.Caption = "EXTREMO"
                                End If
                                If nResultAUX >= CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
'                                If nResultAUX >= Left(psActivacionPCL, 2) And nResultAUX <= Mid(psActivacionPCL, 3, 2) Then JIPR20200824
                                    'lblAct2.Caption = "ACTIVO" JIPR20200824
                                    lblAct2.Caption = "SE ACTIVA PCL"
                                End If
                                
                           Case 3
                                nResultAUX = CCur(rsDatos!nRatioLAMN)
                                If nResultAUX > CCur(psBajo) Then
                                'nResultAUX = rsDatos!nRatioLAMN JIPR20201023
'                                If nResultAUX > psBajo Then
                                    lblRatioLAMN.BackColor = &HC000&
                                    lblRatioLAMN.Caption = "BAJO"
                                End If
                                 If nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                    lblRatioLAMN.BackColor = &HFFFF&
                                    lblRatioLAMN.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblRatioLAMN.BackColor = &H80FF&
                                    lblRatioLAMN.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
'                                If nResultAUX <= psExtremo Then JIPR20201023
                                    lblRatioLAMN.BackColor = &HFF&
                                    lblRatioLAMN.Caption = "EXTREMO"
                                End If
                                
                                If nResultAUX > CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
'                                If nResultAUX > Left(psActivacionPCL, 2) And nResultAUX <= Mid(psActivacionPCL, 3, 2) Then JIPR20201023
                                   'lblAct3.Caption = "ACTIVO" JIPR20200824
                                    lblAct3.Caption = "SE ACTIVA PCL"
                                End If
                                
                           Case 4
                           
                                nResultAUX = CCur(rsDatos!nRatioLAME)
                                If nResultAUX > CCur(psBajo) Then
'                                nResultAUX = rsDatos!nRatioLAME JIPR20201023
'                                If nResultAUX > psBajo Then
                                    lblRatioLAME.BackColor = &HC000&
                                    lblRatioLAME.Caption = "BAJO"
                                End If
                                If nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                               ' If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                    lblRatioLAME.BackColor = &HFFFF&
                                    lblRatioLAME.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblRatioLAME.BackColor = &H80FF&
                                    lblRatioLAME.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
'                                If nResultAUX <= psExtremo Then JIPR20201023
                                    lblRatioLAME.BackColor = &HFF&
                                    lblRatioLAME.Caption = "EXTREMO"
                                End If
                                If nResultAUX > CCur(Left(psActivacionPCL, 2)) And nResultAUX <= CCur(Mid(psActivacionPCL, 3, 2)) Then
'                                If nResultAUX > Left(psActivacionPCL, 2) And nResultAUX <= Mid(psActivacionPCL, 3, 2) Then JIPR20201023
                                     'lblAct4.Caption = "ACTIVO" JIPR20200824
                                    lblAct4.Caption = "SE ACTIVA PCL"
                                End If
                                
                           Case 5
                           
                                nResultAUX = CCur(rsDatos!nRatioILMN)
                                If nResultAUX > CCur(psBajo) Then
'                                nResultAUX = rsDatos!nRatioILMN JIPR20201023
'                                If nResultAUX > psBajo Then
                                    lblRatioILMN.BackColor = &HC000&
                                    lblRatioILMN.Caption = "BAJO"
                                End If
                                 If nResultAUX > CCur(Left(psModerado, 1)) And nResultAUX <= CCur(Right(psModerado, 1)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                    lblRatioILMN.BackColor = &HFFFF&
                                    lblRatioILMN.Caption = "MODERADO"
                                End If
                                 If nResultAUX > CCur(Left(psAlto, 1)) And nResultAUX <= CCur(Right(psAlto, 1)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblRatioILMN.BackColor = &H80FF&
                                    lblRatioILMN.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
                                'If nResultAUX <= psExtremo Then JIPR20201023
                                    lblRatioILMN.BackColor = &HFF&
                                    lblRatioILMN.Caption = "EXTREMO"
                                End If
                           
                           Case 6
                                 nResultAUX = CCur(rsDatos!nRatioCLMN)
                                'nResultAUX = rsDatos!nRatioCLMN JIPR20201023
                                If nResultAUX > CCur(Left(psBajo, 3)) And nResultAUX <= CCur(Right(psBajo, 3)) Then
                                'If nResultAUX > psBajo Then JIPR20201023
                                    lblRatioCLMN.BackColor = &HC000&
                                    lblRatioCLMN.Caption = "BAJO"
                                End If
                                If nResultAUX > CCur(Left(psModerado, 3)) And nResultAUX <= CCur(Right(psModerado, 3)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                     lblRatioCLMN.BackColor = &HFFFF&
                                     lblRatioCLMN.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 3)) And nResultAUX <= CCur(Right(psAlto, 3)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblRatioCLMN.BackColor = &H80FF&
                                    lblRatioCLMN.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
                                'If nResultAUX <= psExtremo Then JIPR20201023
                                    lblRatioCLMN.BackColor = &HFF&
                                    lblRatioCLMN.Caption = "EXTREMO"
                                End If
                                 If nResultAUX > CCur(Left(psActivacionPCL, 3)) And nResultAUX <= CCur(Right(psActivacionPCL, 3)) Then
                                'If nResultAUX <= psActivacionPCL Then JIPR20201023
                                   'lblAct5.Caption = "ACTIVO" JIPR20200824
                                    lblAct5.Caption = "SE ACTIVA PCL"
                                End If
                                
                           Case 7
                                nResultAUX = CCur(rsDatos!nRatioCLME)
                                'nResultAUX = rsDatos!nRatioCLME JIPR20201023
                                If nResultAUX > CCur(Left(psBajo, 3)) And nResultAUX <= CCur(Right(psBajo, 3)) Then
                               ' If nResultAUX > psBajo Then JIPR20201023
                                    lblRatioCLME.BackColor = &HC000&
                                    lblRatioCLME.Caption = "BAJO"
                                End If
                                If nResultAUX > CCur(Left(psModerado, 3)) And nResultAUX <= CCur(Right(psModerado, 3)) Then
                                'If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                    lblRatioCLME.BackColor = &HFFFF&
                                    lblRatioCLME.Caption = "MODERADO"
                                End If
                                If nResultAUX > CCur(Left(psAlto, 3)) And nResultAUX <= CCur(Right(psAlto, 3)) Then
                                'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblRatioCLME.BackColor = &H80FF&
                                    lblRatioCLME.Caption = "ALTO"
                                End If
                                If nResultAUX <= CCur(psExtremo) Then
                                'If nResultAUX <= psExtremo Then JIPR20201023
                                    lblRatioCLME.BackColor = &HFF&
                                    lblRatioCLME.Caption = "EXTREMO"
                                End If
                                If nResultAUX > CCur(Left(psActivacionPCL, 3)) And nResultAUX <= CCur(Right(psActivacionPCL, 3)) Then
                                'If nResultAUX <= psActivacionPCL Then JIPR20201023
                                     'lblAct6.Caption = "ACTIVO" JIPR20200824
                                    lblAct6.Caption = "SE ACTIVA PCL"
                                End If
                                
                          Case 8
                                nResultAUX = CCur(rsDatos!nEncajeMN)
                                If nResultAUX > CCur(Left(psBajo, 2)) And nResultAUX < CCur(Right(psBajo, 2)) Then
                                'nResultAUX = rsDatos!nEncajeMN JIPR20201023
                                'If nResultAUX > Left(psBajo, 2) And nResultAUX < Right(psBajo, 2) Then
                                    lblEncajeMN.BackColor = &HC000&
                                    lblEncajeMN.Caption = "BAJO"
                                End If
                                
                                If (nResultAUX >= CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Mid(psModerado, 3, 2))) Or (nResultAUX >= CCur(Mid(psModerado, 6, 2)) And nResultAUX <= CCur(Mid(psModerado, 8, 2))) Then
                                '(nResultAUX >= Left(psModerado, 2) And nResultAUX <= Mid(psModerado, 3, 2)) Or (nResultAUX >= Mid(psModerado, 6, 2) And nResultAUX <= Mid(psModerado, 8, 2)) Then JIPR20201023
                                    lblEncajeMN.BackColor = &HFFFF&
                                    lblEncajeMN.Caption = "MODERADO"
                                End If
                               If (nResultAUX >= CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Mid(psAlto, 3, 2))) Or (nResultAUX >= CCur(Mid(psAlto, 6, 2)) And nResultAUX <= CCur(Mid(psAlto, 8, 2))) Then
                                'If (nResultAUX >= Left(psAlto, 2) And nResultAUX <= Mid(psAlto, 3, 2)) Or (nResultAUX >= Mid(psAlto, 6, 2) And nResultAUX <= Mid(psAlto, 8, 2)) Then JIPR20201023
                                    lblEncajeMN.BackColor = &H80FF&
                                    lblEncajeMN.Caption = "ALTO"
                                End If
                                
                                'If nResultAUX < Left(psExtremo, 2) And nResultAUX > Right(psExtremo, 2) Then  'Comentado by NAGL 20181006
                                'If nResultAUX < Left(psExtremo, 2) Or nResultAUX > Right(psExtremo, 2) Then 'NAGL 20181006 Según Correo JIPR20201023
                                 If nResultAUX < CCur(Left(psExtremo, 2)) Or nResultAUX > CCur(Right(psExtremo, 2)) Then
                                    lblEncajeMN.BackColor = &HFF&
                                    lblEncajeMN.Caption = "EXTREMO"
                                End If
                            
                          Case 9
                                nResultAUX = CCur(rsDatos!nEncajeME)
                                If nResultAUX > CCur(Left(psBajo, 2)) And nResultAUX < CCur(Right(psBajo, 2)) Then
                                'nResultAUX = rsDatos!nEncajeME JIPR20201023
                                'If nResultAUX > Left(psBajo, 2) And nResultAUX < Right(psBajo, 2) Then
                                    lblEncajeME.BackColor = &HC000&
                                    lblEncajeME.Caption = "BAJO"
                                End If
                                If (nResultAUX >= CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Mid(psModerado, 3, 2))) Or (nResultAUX >= CCur(Mid(psModerado, 6, 2)) And nResultAUX <= CCur(Mid(psModerado, 8, 2))) Then
                                'If (nResultAUX >= Left(psModerado, 2) And nResultAUX <= Mid(psModerado, 3, 2)) Or (nResultAUX >= Mid(psModerado, 6, 2) And nResultAUX <= Mid(psModerado, 8, 2)) Then JIPR20201023
                                    lblEncajeME.BackColor = &HFFFF&
                                    lblEncajeME.Caption = "MODERADO"
                                End If
                                If (nResultAUX >= CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Mid(psAlto, 3, 2))) Or (nResultAUX >= CCur(Mid(psAlto, 6, 2)) And nResultAUX <= CCur(Mid(psAlto, 8, 2))) Then
                                'If (nResultAUX >= Left(psAlto, 2) And nResultAUX <= Mid(psAlto, 3, 2)) Or (nResultAUX >= Mid(psAlto, 6, 2) And nResultAUX <= Mid(psAlto, 8, 2)) Then JIPR20201023
                                    lblEncajeME.BackColor = &H80FF&
                                    lblEncajeME.Caption = "ALTO"
                                End If
                                'If nResultAUX < Left(psExtremo, 2) And nResultAUX > Right(psExtremo, 2) Then 'Comentado by NAGL 20181006
                                'If nResultAUX < Left(psExtremo, 2) Or nResultAUX > Right(psExtremo, 2) Then 'NAGL 20181006 Según Correo JIPR20201023
                                If nResultAUX < CCur(Left(psExtremo, 2)) Or nResultAUX > CCur(Right(psExtremo, 2)) Then
                                  lblEncajeME.BackColor = &HFF&
                                  lblEncajeME.Caption = "EXTREMO"
                                End If
                            
                         Case 12
                                'Agregado by NAGL Según Anx02_ERS006-2019
                                 nResultAUX = CCur(rsDatos!nActivosTOSE)
                                 If nResultAUX > CCur(psBajo) Then
                                 'nResultAUX = rsDatos!nActivosTOSE JIPR20201023
                                 'If nResultAUX > psBajo Then
                                     lblActivosTOSE.BackColor = &HC000&
                                     lblActivosTOSE.Caption = "BAJO"
                                 End If
                                 If nResultAUX > CCur(Left(psModerado, 2)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                ' If nResultAUX > psModerado And nResultAUX <= psBajo Then JIPR20201023
                                    lblActivosTOSE.BackColor = &HFFFF&
                                    lblActivosTOSE.Caption = "MODERADO"
                                 End If
                                 If nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 2)) Then
                                 'If nResultAUX > psAlto And nResultAUX <= psModerado Then JIPR20201023
                                    lblActivosTOSE.BackColor = &H80FF&
                                    lblActivosTOSE.Caption = "ALTO"
                                 End If
                                 If nResultAUX <= CCur(psExtremo) Then
                                'If nResultAUX <= psExtremo Then
                                    lblActivosTOSE.BackColor = &HFF&
                                    lblActivosTOSE.Caption = "EXTREMO"
                                 End If
                            Case 13
                                'JIPR20201023
                                 nResultAUX = CCur(rsDatos!nActivosTotales)
                                 If nResultAUX > CCur(psBajo) Then
                                     lblActivoTotales.BackColor = &HC000&
                                     lblActivoTotales.Caption = "BAJO"
                                 End If
                                 If nResultAUX > CCur(Left(psModerado, 4)) And nResultAUX <= CCur(Right(psModerado, 2)) Then
                                    lblActivoTotales.BackColor = &HFFFF&
                                    lblActivoTotales.Caption = "MODERADO"
                                 End If
                                 If nResultAUX > CCur(Left(psAlto, 2)) And nResultAUX <= CCur(Right(psAlto, 4)) Then
                                    lblActivoTotales.BackColor = &H80FF&
                                    lblActivoTotales.Caption = "ALTO"
                                 End If
                                 If nResultAUX > CCur(Left(psExtremo, 4)) And nResultAUX < CCur(Right(psExtremo, 2)) Then
                                    lblActivoTotales.BackColor = &HFF&
                                    lblActivoTotales.Caption = "EXTREMO"
                                 End If
                                 'JIPR20201023
                     End Select
                     
                    rsNiveles.MoveNext
                    Loop
        Else
        'MsgBox "No se encontraron los Anexos 15A y 15B con la fecha ingresada...!!" & Fecha, vbInformation, "Aviso"
         MsgBox "Por favor asegúrese de haber generado los dos Anexos 15A y 15B con la fecha Ingresada." & Fecha, vbInformation, "Aviso"
        End If
End Sub
Public Function SeleccionaAlertasTempranA(ByVal pdFecha As Date) As ADODB.Recordset
On Error GoTo SeleccionaAlertasTempranAErr
   Dim oRS As ADODB.Recordset
   Dim oConec As DConecta
   Dim psSql As String
   Set oRS = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_Alertastempranas '" & Format(pdFecha, "YYYY/MM/DD") & "'"
   Set oRS = oConec.CargaRecordSet(psSql)
   Set SeleccionaAlertasTempranA = oRS
   oConec.CierraConexion
Exit Function
SeleccionaAlertasTempranAErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Function


Public Function SeleccionaNiveles() As ADODB.Recordset
On Error GoTo SeleccionaNivelesErr
   Dim oRS As ADODB.Recordset
   Dim oConec As DConecta
   Dim psSql As String
   Set oRS = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_NivelesAlertas "
   Set oRS = oConec.CargaRecordSet(psSql)
   Set SeleccionaNiveles = oRS
   oConec.CierraConexion
Exit Function
SeleccionaNivelesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Function
Private Sub optTipoBus_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFecha.Visible Then
        txtFecha.SetFocus
        
    End If
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmProcesar.SetFocus
End If
End Sub
Private Sub txtFecha_GotFocus()
With txtFecha
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


