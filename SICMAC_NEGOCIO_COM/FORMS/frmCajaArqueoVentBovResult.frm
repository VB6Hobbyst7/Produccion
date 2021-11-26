VERSION 5.00
Begin VB.Form frmCajaArqueoVentBovResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueo: Resultado "
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "frmCajaArqueoVentBovResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Resultado Dólares "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   3480
      TabIndex        =   11
      Top             =   600
      Width           =   3255
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Sistema :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblMontoDolSist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1695
         TabIndex        =   16
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Billetaje :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label lblMontoDolBill 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1695
         TabIndex        =   14
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label lblEtiqDol 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Faltante/Sobrante :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label lblFaltSobrDol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1695
         TabIndex        =   12
         Top             =   855
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Resultado Soles "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3255
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Sistema :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblMontoSolSist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Billetaje :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label lblMontoSolBill 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label lblEtiqSol 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Faltante/Sobrante :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label lblFaltSobrSol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   855
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   1
      Top             =   2070
      Width           =   1170
   End
   Begin VB.CommandButton cmdImprActa 
      Caption         =   "Imprimir Acta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   2070
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultado del Arqueo :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   285
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1755
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   6660
   End
End
Attribute VB_Name = "frmCajaArqueoVentBovResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmOpeArqueoVentBovResult
'** Descripción : Resultado del proceso de Arqueos de Ventanillas y Bóvedas creado segun RFC081-2012
'** Creación : JUEZ, 20120808 06:00:00 PM
'********************************************************************

Option Explicit
Dim ImprActa As Boolean
Dim MatDat As Variant
Dim MatPrs As Variant
Dim lnTipoArq As Integer
Dim prsBillMN As ADODB.Recordset, prsMonMN As ADODB.Recordset
Dim prsBillME As ADODB.Recordset

Private Sub cmdImprActa_Click()
    Call ImprimeActa1
    Call ImprimeActa2
End Sub

Private Sub cmdsalir_Click()
    If ImprActa = False Then
        If MsgBox("Desea salir sin antes imprimir las actas?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Else
        If MsgBox("Desea salir del resultado del Arqueo?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Public Sub Inicio(ByVal MatDatos As Variant, ByVal MatPersRel As Variant, ByVal pnTipoArq As Integer, ByVal pbConforme As Boolean, _
                  ByVal pnResultSolSist As Double, ByVal pnResultSolBill As Double, ByVal pnResultDolSist As Double, _
                  ByVal pnResultDolBill As Double, ByVal rsBillMN As ADODB.Recordset, _
                  ByVal rsMonMN As ADODB.Recordset, ByVal rsBillME As ADODB.Recordset)
    
    Me.Caption = Me.Caption + IIf(pnTipoArq = 1, "Ventanilla", IIf(pnTipoArq = 3, "Entre Ventanillas", "Bóveda"))
    
    lblResult.Caption = IIf(pbConforme, "CONFORME", "NO CONFORME")
    lblResult.ForeColor = IIf(pbConforme, &H800000, &HFF&)
    lblMontoSolBill.Caption = Format(pnResultSolBill, "#,##0.00")
    lblMontoSolSist.Caption = Format(pnResultSolSist, "#,##0.00")
    lblMontoDolBill.Caption = Format(pnResultDolBill, "#,##0.00")
    lblMontoDolSist.Caption = Format(pnResultDolSist, "#,##0.00")
    
    lblFaltSobrSol.Caption = Format(pnResultSolBill - pnResultSolSist, "#,##0.00")
    lblFaltSobrDol.Caption = Format(pnResultDolBill - pnResultDolSist, "#,##0.00")
    
    If CDbl(lblFaltSobrSol.Caption) < 0 Then
        lblFaltSobrSol.ForeColor = &HFF&
        lblEtiqSol.ForeColor = &HFF&
    End If
    If CDbl(lblFaltSobrDol.Caption) < 0 Then
        lblFaltSobrDol.ForeColor = &HFF&
        lblEtiqDol.ForeColor = &HFF&
    End If
    
    ImprActa = False
    
    lnTipoArq = pnTipoArq
    MatDat = MatDatos
    MatPrs = MatPersRel
    
    Set prsBillMN = rsBillMN
    Set prsMonMN = rsMonMN
    Set prsBillME = rsBillME
    
    Me.Show 1
End Sub

Private Sub ImprimeActa1()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim FilaPertenece As Integer
    Dim TotBillMN As Double
    Dim TotMonMN As Double
    Dim TotBillME As Double
    Dim i As Integer
    
    On Error GoTo ErrorImprActa
    
    'If rs.RecordCount > 0 Then
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        
        If lnTipoArq = 1 Or lnTipoArq = 3 Then 'RIRO20140719 ER072 Add "lnTipoArq = 3"
            lsNomHoja = "Hoja1"
            lsFile = "ActaArqueoVentanilla01"
        Else
            lsNomHoja = "Hoja1"
            lsFile = "ActaArqueoCajaGralMN"
        End If
        
        lsArchivo = "\spooler\" & lsFile & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & ".xls"
        If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
            Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
        Else
            MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
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
        
        If lnTipoArq = 1 Or lnTipoArq = 3 Then 'RIRO20140719 ERS072 Add lnTipoArq = 3
            xlHoja1.Cells(6, 1) = MatDat(0, 1)
            xlHoja1.Cells(6, 4) = Format(MatDat(0, 2), "dd/mm/yyyy")
            xlHoja1.Cells(6, 7) = Format(Right(MatDat(0, 2), 8), "hh:mm:ss")
            
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 2 Then
                    xlHoja1.Cells(9, 1) = MatPrs(i, 0)
                    xlHoja1.Cells(9, 2) = MatPrs(i, 2)
                    Exit For
                End If
            Next i
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 1 Then
                    xlHoja1.Cells(12, 1) = MatPrs(i, 0)
                    xlHoja1.Cells(12, 2) = MatPrs(i, 2)
                    Exit For
                End If
            Next i

            If prsBillMN.RecordCount > 0 Then 'Billetes MN
                prsBillMN.MoveFirst
                For i = 0 To prsBillMN.RecordCount - 1
                    Select Case Trim(prsBillMN!Denominación)
                    Case "200.00"
                        FilaPertenece = 17
                    Case "100.00"
                        FilaPertenece = 18
                    Case "50.00"
                        FilaPertenece = 19
                    Case "20.00"
                        FilaPertenece = 20
                    Case "10.00"
                        FilaPertenece = 21
                    End Select
                    xlHoja1.Cells(FilaPertenece, 2) = prsBillMN!Cant
                    xlHoja1.Cells(FilaPertenece, 3) = Format(prsBillMN!Monto, "#,##0.00")
                    TotBillMN = Format(TotBillMN + prsBillMN!Monto, "#,##0.00")
                    prsBillMN.MoveNext
                Next i
                xlHoja1.Cells(22, 3) = Format(TotBillMN, "#,##0.00")
            End If

            If prsMonMN.RecordCount > 0 Then 'Monedas MN
                prsMonMN.MoveFirst
                For i = 0 To prsMonMN.RecordCount - 1
                    Select Case Trim(prsMonMN!Denominación)
                    Case "5.00"
                        FilaPertenece = 27
                    Case "2.00"
                        FilaPertenece = 28
                    Case "1.00"
                        FilaPertenece = 29
                    Case "0.50"
                        FilaPertenece = 30
                    Case "0.20"
                        FilaPertenece = 32
                    Case "0.10"
                        FilaPertenece = 33
                    Case "0.05"
                        FilaPertenece = 34
                    Case "0.01"
                        FilaPertenece = 35
                    End Select
                    xlHoja1.Cells(FilaPertenece, 2) = prsMonMN!Cant
                    xlHoja1.Cells(FilaPertenece, 3) = Format(prsMonMN!Monto, "#,##0.00")
                    TotMonMN = Format(TotMonMN + prsMonMN!Monto, "#,##0.00")
                    prsMonMN.MoveNext
                Next i
                xlHoja1.Cells(37, 3) = Format(TotMonMN, "#,##0.00")
            End If
            
            xlHoja1.Cells(38, 3) = Format(TotBillMN + TotMonMN, "#,##0.00")
            xlHoja1.Cells(40, 3) = Format(CDbl(lblMontoSolSist.Caption), "#,##0.00")
            
            xlHoja1.Cells(42, 3) = (TotBillMN + TotMonMN) - CDbl(lblMontoSolSist.Caption)
            
            If prsBillME.RecordCount > 0 Then 'Billetes ME
                prsBillME.MoveFirst
                For i = 0 To prsBillME.RecordCount - 1
                    Select Case Trim(prsBillME!Denominación)
                    Case "100.00"
                        FilaPertenece = 17
                    Case "50.00"
                        FilaPertenece = 18
                    Case "20.00"
                        FilaPertenece = 19
                    Case "10.00"
                        FilaPertenece = 20
                    Case "5.00"
                        FilaPertenece = 21
                    Case "2.00"
                        FilaPertenece = 22
                    Case "1.00"
                        FilaPertenece = 23
                    End Select
                    xlHoja1.Cells(FilaPertenece, 7) = prsBillME!Cant
                    xlHoja1.Cells(FilaPertenece, 9) = Format(prsBillME!Monto, "#,##0.00")
                    TotBillME = Format(TotBillME + prsBillME!Monto, "#,##0.00")
                    prsBillME.MoveNext
                Next i
                xlHoja1.Cells(27, 9) = Format(TotBillME, "#,##0.00")
            End If

            xlHoja1.Cells(28, 9) = Format(CDbl(lblMontoDolSist.Caption), "#,##0.00")
            xlHoja1.Cells(29, 9) = Format(TotBillME - CDbl(lblMontoDolSist.Caption), "#,##0.00")
        Else
            xlHoja1.Cells(6, 4) = MatDat(0, 1)
            xlHoja1.Cells(6, 7) = Format(Right(MatDat(0, 4), 8), "hh:mm:ss")
            xlHoja1.Cells(6, 9) = Format(MatDat(0, 4), "dd/mm/yyyy")
            
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 2 Then
                    xlHoja1.Cells(7, 6) = Replace(Replace(MatPrs(i, 2), "/", " "), ",", " ")
                    Exit For
                End If
            Next i
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 1 Then
                    xlHoja1.Cells(8, 4) = Replace(Replace(MatPrs(i, 2), "/", " "), ",", " ")
                    'xlHoja1.Cells(8, 9) = Replace(oUser.PersCargo, " DE CONTABILIDAD", "")
                    Exit For
                End If
            Next i
            
'            xlHoja1.Cells(7, 8) = "" 'responsable
'            xlHoja1.Cells(8, 4) = "" 'persona de contabilidad
'            xlHoja1.Cells(8, 9) = "" 'puesto de contabilidad
            
            If prsBillMN.RecordCount > 0 Then 'Billetes MN
                prsBillMN.MoveFirst
                For i = 0 To prsBillMN.RecordCount - 1
                    Select Case Trim(prsBillMN!Denominación)
                    Case "200.00"
                        FilaPertenece = 15
                    Case "100.00"
                        FilaPertenece = 16
                    Case "50.00"
                        FilaPertenece = 17
                    Case "20.00"
                        FilaPertenece = 18
                    Case "10.00"
                        FilaPertenece = 19
                    End Select
                    xlHoja1.Cells(FilaPertenece, 4) = prsBillMN!Cant
                    xlHoja1.Cells(FilaPertenece, 7) = Format(prsBillMN!Monto, "#,##0.00")
                    TotBillMN = Format(TotBillMN + prsBillMN!Monto, "#,##0.00")
                    prsBillMN.MoveNext
                Next i
                xlHoja1.Cells(19, 10) = Format(TotBillMN, "#,##0.00")
            End If

            If prsMonMN.RecordCount > 0 Then 'Monedas MN
                prsMonMN.MoveFirst
                For i = 0 To prsMonMN.RecordCount - 1
                    Select Case Trim(prsMonMN!Denominación)
                    Case "5.00"
                        FilaPertenece = 24
                    Case "2.00"
                        FilaPertenece = 25
                    Case "1.00"
                        FilaPertenece = 26
                    Case "0.50"
                        FilaPertenece = 27
                    Case "0.20"
                        FilaPertenece = 28
                    Case "0.10"
                        FilaPertenece = 29
                    Case "0.05"
                        FilaPertenece = 30
                    Case "0.01"
                        FilaPertenece = 31
                    End Select
                    xlHoja1.Cells(FilaPertenece, 4) = prsMonMN!Cant
                    xlHoja1.Cells(FilaPertenece, 7) = Format(prsMonMN!Monto, "#,##0.00")
                    TotMonMN = Format(TotMonMN + prsMonMN!Monto, "#,##0.00")
                    prsMonMN.MoveNext
                Next i
                xlHoja1.Cells(31, 10) = Format(TotMonMN, "#,##0.00")
            End If
            
            xlHoja1.Cells(32, 10) = Format(TotBillMN + TotMonMN, "#,##0.00")
            
            xlHoja1.Cells(58, 10) = xlHoja1.Cells(32, 10) + xlHoja1.Cells(43, 10) + xlHoja1.Cells(54, 10)
            
            xlHoja1.Cells(60, 10) = lblFaltSobrSol.Caption
            
            xlHoja1.Cells(63, 2) = Format(Right(MatDat(0, 2), 8), "hh:mm:ss")
            xlHoja1.Cells(63, 4) = Format(MatDat(0, 2), "dd/mm/yyyy")
        End If
        
        Dim psArchivoAGrabarC As String
        
        xlHoja1.SaveAs App.path & lsArchivo
        psArchivoAGrabarC = App.path & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Acta Generada Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
        ImprActa = True
    'Else
    '    MsgBox "No hay Datos", vbInformation, "Aviso"
    'End If
    
    Exit Sub
ErrorImprActa:
    ImprActa = False
    MsgBox err.Description, vbInformation, "Error!!"
End Sub

Private Sub ImprimeActa2()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim Fila As Integer
    Dim Columna As Integer
    Dim lsFiltroGrupoOpe As String
    Dim TotIng As Double
    Dim TotEgre As Double
    Dim i As Integer, j As Integer, FilaPertenece As Integer
    Dim psUserPersArqueado As String
    Dim TotBillME As Double
    Dim rs As ADODB.Recordset
    
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    
    On Error GoTo ErrorImprActa
    
    'If rs.RecordCount > 0 Then
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        
        If lnTipoArq = 1 Or lnTipoArq = 3 Then 'RIRO20140719 ERS072
            lsNomHoja = "Hoja1"
            lsFile = "ActaArqueoVentanilla02"
        Else
            lsNomHoja = "Hoja1"
            lsFile = "ActaArqueoCajaGralME"
        End If
        
        lsArchivo = "\spooler\" & lsFile & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & ".xls"
        If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
            Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
        Else
            MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
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
        
        If lnTipoArq = 1 Or lnTipoArq = 3 Then 'RIRO20140719 ERS072
            xlHoja1.Cells(6, 1) = MatDat(0, 1)
            xlHoja1.Cells(6, 4) = Format(MatDat(0, 2), "dd/mm/yyyy")
            xlHoja1.Cells(6, 7) = Format(Right(MatDat(0, 2), 8), "hh:mm:ss")
            
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 2 Then
                    xlHoja1.Cells(9, 1) = MatPrs(i, 0)
                    xlHoja1.Cells(9, 2) = MatPrs(i, 2)
                    psUserPersArqueado = MatPrs(i, 0)
                    Exit For
                End If
            Next i
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 1 Then
                    xlHoja1.Cells(12, 1) = MatPrs(i, 0)
                    xlHoja1.Cells(12, 2) = MatPrs(i, 2)
                    Exit For
                End If
            Next i

            xlHoja1.Cells(15, 6) = Format(MatDat(0, 2), "yyyy")
            xlHoja1.Cells(15, 7) = Format(MatDat(0, 2), "mm")
            xlHoja1.Cells(15, 9) = Format(MatDat(0, 2), "dd")
            
            For i = 1 To 2
                Columna = IIf(i = 1, 6, 8)
                Set rs = oCajero.GetOpeIngEgreCaptacionesDetalle(gGruposIngEgreAgeLocal, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(19, Columna) = TotIng
                xlHoja1.Cells(31, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreAgeLocal, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 1)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(17, Columna) = TotIng
                xlHoja1.Cells(29, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreAgeLocal, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 2)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(18, Columna) = TotIng
                xlHoja1.Cells(30, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreAgeLocal, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 3)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(20, Columna) = TotIng
                'xlHoja1.Cells(30, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                Set rs = oCajero.GetOpeIngEgreCaptacionesDetalle(gGruposIngEgreOtraAgencia, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(19, Columna) = xlHoja1.Cells(19, Columna) + TotIng
                xlHoja1.Cells(31, Columna) = xlHoja1.Cells(31, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreOtraAgencia, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 1)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(17, Columna) = xlHoja1.Cells(17, Columna) + TotIng
                xlHoja1.Cells(29, Columna) = xlHoja1.Cells(29, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreOtraAgencia, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 2)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(18, Columna) = xlHoja1.Cells(18, Columna) + TotIng
                xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreColocacionesDetalle(gGruposIngEgreOtraAgencia, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC, 3)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(20, Columna) = xlHoja1.Cells(20, Columna) + TotIng
                'xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna)+TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                '******************************
                Set rs = oCajero.GetOpeIngEgreOpeCMACsDetalle(gGruposIngEgreOtraCMAC, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 1)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(17, Columna) = xlHoja1.Cells(17, Columna) + TotIng
                xlHoja1.Cells(29, Columna) = xlHoja1.Cells(29, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOpeCMACsDetalle(gGruposIngEgreOtraCMAC, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 2)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(18, Columna) = xlHoja1.Cells(18, Columna) + TotIng
                xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOpeCMACsDetalle(gGruposIngEgreOtraCMAC, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 3)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(20, Columna) = xlHoja1.Cells(20, Columna) + TotIng
                'xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna)+TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOpeCMACsDetalle(gGruposIngEgreOtraCMAC, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 4)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(19, Columna) = xlHoja1.Cells(19, Columna) + TotIng
                xlHoja1.Cells(31, Columna) = xlHoja1.Cells(31, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                '******************************
                Set rs = oCajero.GetOpeIngEgreOtrasOpeDetalle(gGruposIngEgreOtrasOpe, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 1)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(17, Columna) = xlHoja1.Cells(17, Columna) + TotIng
                xlHoja1.Cells(29, Columna) = xlHoja1.Cells(29, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOtrasOpeDetalle(gGruposIngEgreOtrasOpe, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 2)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(18, Columna) = xlHoja1.Cells(18, Columna) + TotIng
                xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOtrasOpeDetalle(gGruposIngEgreOtrasOpe, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 3)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(20, Columna) = xlHoja1.Cells(20, Columna) + TotIng
                'xlHoja1.Cells(30, Columna) = xlHoja1.Cells(30, Columna)+TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreOtrasOpeDetalle(gGruposIngEgreOtrasOpe, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 4)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(19, Columna) = xlHoja1.Cells(19, Columna) + TotIng
                xlHoja1.Cells(31, Columna) = xlHoja1.Cells(31, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                '******************************
                Set rs = oCajero.GetOpeIngEgreServiciosDetalle(gGruposIngEgreServicios, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 1)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(23, Columna) = TotIng
                xlHoja1.Cells(34, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                Set rs = oCajero.GetOpeIngEgreServiciosDetalle(gGruposIngEgreServicios, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser, 2)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(27, Columna) = TotIng
                xlHoja1.Cells(39, Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                
                ' *** Modificado Por RIRO 20130917
                Set rs = oCajero.GetOpeIngEgreRecaudo(CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(27, Columna) = xlHoja1.Cells(27, Columna) + TotIng
                xlHoja1.Cells(36, Columna) = xlHoja1.Cells(36, Columna) + TotEgre
                TotIng = 0
                TotEgre = 0
                ' *** Fin RIRO
                
                Set rs = oCajero.GetOpeIngEgreCompraVentaDetalle(gGruposIngEgreCompraVenta, CDate(gdFecSis), MatDat(0, 3), psUserPersArqueado, i, gdFecSis, gsCodUser)
                For j = 0 To rs.RecordCount - 1
                    If rs!Efectivo > 0 Then
                        TotIng = TotIng + rs!Efectivo
                    Else
                        TotEgre = TotEgre + rs!Efectivo
                    End If
                    rs.MoveNext
                Next j
                xlHoja1.Cells(IIf(i = 1, 21, 22), Columna) = TotIng
                xlHoja1.Cells(IIf(i = 1, 32, 33), Columna) = TotEgre
                TotIng = 0
                TotEgre = 0
                '******************************
                'Set rs = oCajero.GetOpeIngEgreHabDevDetalle(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, i, gdFecSis, gsCodUser, 3)
                'For J = 0 To rs.RecordCount - 1
                '    nEfec = nEfec + rs!Efectivo
                '    rs.MoveNext
                'Next J
                'Set rs = oCajero.GetOpeIngEgreHabDevDetalle(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, i, gdFecSis, gsCodUser, 3)
                'For J = 0 To rs.RecordCount - 1
                '    nEfec = nEfec + rs!Efectivo
                '    rs.MoveNext
                'Next J
                Set rs = oCajero.GetOpeIngEgreHabDevCajeroDetalle(gGruposIngEgreHabDev, CDate(gdFecSis), MatDat(0, 3), gsCodUser, i, gdFecSis, gsCodUser)
                For j = 0 To rs.RecordCount - 1
                    If rs!cOpecod = "901017" Then 'Habilitacion
                        If rs!Efectivo > 0 Then
                            xlHoja1.Cells(25, Columna) = xlHoja1.Cells(25, Columna) + rs!Efectivo
                        Else
                            xlHoja1.Cells(25, Columna) = Format(0, "#,##0.00")
                        End If
                    ElseIf rs!cOpecod = "901013" Then 'Transferencia
                        If rs!Efectivo > 0 Then
                            xlHoja1.Cells(26, Columna) = xlHoja1.Cells(26, Columna) + rs!Efectivo
                        ElseIf rs!Efectivo < 0 Then
                            xlHoja1.Cells(38, Columna) = xlHoja1.Cells(38, Columna) + rs!Efectivo
                        Else
                            xlHoja1.Cells(26, Columna) = Format(0, "#,##0.00")
                            xlHoja1.Cells(38, Columna) = Format(0, "#,##0.00")
                        End If
                    End If
                    rs.MoveNext
                Next j
                TotIng = 0
                TotEgre = 0
            Next i
        Else
            xlHoja1.Cells(6, 4) = MatDat(0, 1)
            xlHoja1.Cells(6, 8) = Format(Right(MatDat(0, 4), 8), "hh:mm:ss")
            xlHoja1.Cells(6, 10) = Format(MatDat(0, 4), "dd/mm/yyyy")
            
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 2 Then
                    xlHoja1.Cells(7, 8) = Replace(Replace(MatPrs(i, 2), "/", " "), ",", " ")
                    Exit For
                End If
            Next i
            For i = 0 To UBound(MatPrs) - 1
                If MatPrs(i, 1) = 1 Then
                    xlHoja1.Cells(8, 5) = Replace(Replace(MatPrs(i, 2), "/", " "), ",", " ")
                    'xlHoja1.Cells(8, 9) = Replace(oUser.PersCargo, " DE CONTABILIDAD", "")
                    Exit For
                End If
            Next i
            
'            xlHoja1.Cells(7, 8) = "" 'responsable
'            xlHoja1.Cells(8, 4) = "" 'persona de contabilidad
'            xlHoja1.Cells(8, 9) = "" 'puesto de contabilidad
            
            If prsBillME.RecordCount > 0 Then 'Billetes ME
                prsBillME.MoveFirst
                For i = 0 To prsBillME.RecordCount - 1
                    Select Case Trim(prsBillME!Denominación)
                    Case "100.00"
                        FilaPertenece = 15
                    Case "50.00"
                        FilaPertenece = 16
                    Case "20.00"
                        FilaPertenece = 17
                    Case "10.00"
                        FilaPertenece = 18
                    Case "5.00"
                        FilaPertenece = 19
                    Case "2.00"
                        FilaPertenece = 20
                    Case "1.00"
                        FilaPertenece = 21
                    End Select
                    xlHoja1.Cells(FilaPertenece, 4) = prsBillME!Cant
                    xlHoja1.Cells(FilaPertenece, 6) = Format(prsBillME!Monto, "#,##0.00")
                    TotBillME = Format(TotBillME + prsBillME!Monto, "#,##0.00")
                    prsBillME.MoveNext
                Next i
                xlHoja1.Cells(22, 9) = Format(TotBillME, "#,##0.00")
                
                xlHoja1.Cells(49, 9) = xlHoja1.Cells(22, 9)
                xlHoja1.Cells(50, 9) = lblFaltSobrDol.Caption
                
                xlHoja1.Cells(54, 2) = Format(Right(MatDat(0, 2), 8), "hh:mm:ss")
                xlHoja1.Cells(54, 4) = Format(MatDat(0, 2), "dd/mm/yyyy")
            End If
        End If
        
        Dim psArchivoAGrabarC As String
        
        xlHoja1.SaveAs App.path & lsArchivo
        psArchivoAGrabarC = App.path & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Acta Generada Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
        ImprActa = True
    'Else
    '    MsgBox "No hay Datos", vbInformation, "Aviso"
    'End If
    
    Exit Sub
ErrorImprActa:
    ImprActa = False
    MsgBox err.Description, vbInformation, "Error!!"
End Sub
