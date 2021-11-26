VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRRHHRepGen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "frmRRHHRepGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIzqU 
      Caption         =   "&<"
      Height          =   375
      Left            =   3998
      TabIndex        =   5
      Top             =   3720
      Width           =   405
   End
   Begin VB.CommandButton cmdIzqT 
      Caption         =   "<<"
      Height          =   345
      Left            =   3998
      TabIndex        =   4
      Top             =   3390
      Width           =   405
   End
   Begin VB.CommandButton cmdDerT 
      Caption         =   ">>"
      Height          =   375
      Left            =   3998
      TabIndex        =   3
      Top             =   3030
      Width           =   405
   End
   Begin VB.CommandButton cmdDerU 
      Caption         =   "&>"
      Height          =   375
      Left            =   3998
      TabIndex        =   2
      Top             =   2670
      Width           =   405
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   390
      Left            =   6270
      TabIndex        =   13
      Top             =   4065
      Width           =   990
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   3705
      TabIndex        =   10
      Top             =   2160
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7335
      TabIndex        =   9
      Top             =   4065
      Width           =   990
   End
   Begin VB.ListBox lstA 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   5190
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   165
      Width           =   3135
   End
   Begin VB.ListBox lstP 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   45
      TabIndex        =   0
      Top             =   165
      Width           =   3135
   End
   Begin VB.Frame fraTexto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Texto"
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
      Height          =   1035
      Left            =   3210
      TabIndex        =   11
      Top             =   75
      Width           =   1935
      Begin VB.OptionButton optTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Texto"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   990
         TabIndex        =   18
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optLista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Lista"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   255
         Width           =   885
      End
      Begin VB.TextBox txtTexto 
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   615
         Width           =   1740
      End
      Begin VB.ComboBox cmbLista 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   615
         Width           =   1740
      End
   End
   Begin VB.Frame fraNumeros 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Numero"
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
      Height          =   1035
      Left            =   3210
      TabIndex        =   19
      Top             =   75
      Width           =   1935
      Begin VB.TextBox txtHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   795
         TabIndex        =   23
         Top             =   645
         Width           =   1050
      End
      Begin VB.TextBox txtDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   795
         TabIndex        =   22
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Desde :"
         Height          =   180
         Left            =   105
         TabIndex        =   21
         Top             =   307
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta :"
         Height          =   225
         Left            =   135
         TabIndex        =   20
         Top             =   705
         Width           =   555
      End
   End
   Begin VB.Frame fraFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
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
      Height          =   1035
      Left            =   3210
      TabIndex        =   1
      Top             =   75
      Width           =   1935
      Begin MSMask.MaskEdBox mskFin 
         Height          =   315
         Left            =   750
         TabIndex        =   7
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   315
         Left            =   750
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   705
         Width           =   555
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Desde :"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   307
         Width           =   555
      End
   End
   Begin VB.Frame fraOperador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   3210
      TabIndex        =   25
      Top             =   1020
      Width           =   1935
      Begin VB.OptionButton optDiff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Diferente"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   885
         TabIndex        =   27
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton optIgual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Igual"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   105
         TabIndex        =   26
         Top             =   210
         Value           =   -1  'True
         Width           =   720
      End
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   315
      Left            =   3675
      SizeMode        =   1  'Stretch
      TabIndex        =   24
      Top             =   4155
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmRRHHRepGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmdAplicar_Click()
    Dim lnPos As Integer
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    
    If lstA.ListIndex = -1 Then Exit Sub
    
    lnPos = InStr(1, Me.lstA.List(lstA.ListIndex), "*", vbTextCompare)
    lnPosIni = InStr(1, Me.lstA.List(lstA.ListIndex), ",", vbTextCompare) + 1
    lnPosFin = InStr(1, Me.lstA.List(lstA.ListIndex), "-", vbTextCompare)
    
    If Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "datetime" Then
        If Not IsDate(Me.mskIni.Text) Then
            MsgBox "Ingrese una fecha valida.", vbInformation, "Aviso"
            Me.mskIni.SetFocus
            Exit Sub
        ElseIf Not IsDate(Me.mskFin.Text) Then
            MsgBox "Ingrese una fecha valida.", vbInformation, "Aviso"
            Me.mskFin.SetFocus
            Exit Sub
        Else
            If lnPos = 0 Then
                If optIgual.value Then
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & Me.mskIni.Text & "$" & Me.mskFin.Text & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & Me.mskIni.Text & "$" & Me.mskFin.Text & "{>}"
                End If
            Else
                If optIgual.value Then
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & Me.mskIni.Text & "$" & Me.mskFin.Text & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & Me.mskIni.Text & "$" & Me.mskFin.Text & "{>}"
                End If
            End If
        End If
    ElseIf Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "char" Or Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "varchar" Then
        If lnPos = 0 Then
            If Me.optLista.value Then
                If Me.cmbLista.Text <> "" Then
                    If Me.optIgual.value Then
                        lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & "1" & Me.cmbLista.Text & "{=}"
                    Else
                        lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & "1" & Me.cmbLista.Text & "{>}"
                    End If
                End If
            Else
                If Me.optIgual.value Then
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & "0" & Trim(Me.txtTexto.Text) & "%" & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & "0" & Trim(Me.txtTexto.Text) & "%" & "{>}"
                End If
            End If
        Else
            If Me.optLista.value Then
                If Me.cmbLista.Text = "" Then
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos - 1)
                Else
                    If Me.optIgual.value Then
                        lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & "1" & Me.cmbLista.Text & "{=}"
                    Else
                        lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & "1" & Me.cmbLista.Text & "{>}"
                    End If
                End If
            Else
                If Me.optIgual.value Then
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & "0" & Trim(Me.txtTexto.Text) & "%" & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & "0" & Trim(Me.txtTexto.Text) & "%" & "{>}"
                End If
            End If
        End If
        Me.fraFechas.Visible = False
    ElseIf Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "numeric" Or Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "money" Then
        If Not IsNumeric(Me.txtDesde.Text) Then
            MsgBox "Ingrese un numero valido.", vbInformation, "Aviso"
            Me.txtDesde.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Me.txtHasta.Text) Then
            MsgBox "Ingrese una numero valido.", vbInformation, "Aviso"
            Me.txtDesde.SetFocus
            Exit Sub
        Else
            If lnPos = 0 Then
                If optIgual.value Then
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & Me.txtDesde.Text & "$" & Me.txtHasta.Text & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = lstA.List(lstA.ListIndex) & "*" & Me.txtDesde.Text & "$" & Me.txtHasta.Text & "{>}"
                End If
            Else
                If optIgual.value Then
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & Me.txtDesde.Text & "$" & Me.txtHasta.Text & "{=}"
                Else
                    lstA.List(lstA.ListIndex) = Left(lstA.List(lstA.ListIndex), lnPos) & Me.txtDesde.Text & "$" & Me.txtHasta.Text & "{>}"
                End If
            End If
            Me.fraFechas.Visible = False
        End If
    End If
End Sub

Private Sub cmdDerT_Click()
    Dim i As Integer
    For i = 0 To Me.lstP.ListCount - 1
        Me.lstA.AddItem lstP.List(i)
        lstA.Selected(lstA.ListCount - 1) = True
    Next i
End Sub

Private Sub cmdDerU_Click()
    If Me.lstP.ListIndex <> -1 Then
        Me.lstA.AddItem lstP.List(lstP.ListIndex)
        lstA.Selected(lstA.ListCount - 1) = True
    End If
End Sub

Private Sub cmdIzqT_Click()
    lstA.Clear
End Sub

Private Sub cmdIzqU_Click()
    Dim lnAux As Integer
    If Me.lstA.ListIndex <> -1 Then
        lnAux = lstA.ListIndex
        Me.lstA.RemoveItem lstA.ListIndex
        If lstA.ListCount > lnAux Then
            lstA.ListIndex = lnAux
        Else
            lstA.ListIndex = lstA.ListCount - 1
        End If
    End If
End Sub

Private Sub cmdReporte_Click()
    Dim sqlP As String
    Dim sqlAux As String
    Dim sqlFilto As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim i As Integer
    Dim J As Integer
    Dim lnPos As Integer
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnPosDiff As Integer
    Dim lnPosOtro As Integer
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lbBanFiltro As Boolean
    Dim oRep As DRHReportes
    
    Set oRep = New DRHReportes
    Set rsP = New ADODB.Recordset
    Dim lnI As Integer
    
    
    
    If Me.lstA.ListCount = 0 Then Exit Sub
    
    lbBanFiltro = True
    For i = 0 To Me.lstA.ListCount - 1
        lnPos = InStr(1, Me.lstA.List(i), "*", vbTextCompare)
        lnPosIni = InStr(1, Me.lstA.List(i), ",", vbTextCompare) + 1
        lnPosFin = InStr(1, Me.lstA.List(i), "-", vbTextCompare)
        lnPosDiff = InStr(1, Me.lstA.List(i), "{", vbTextCompare)
        lnPosOtro = InStr(1, Me.lstA.List(i), "$", vbTextCompare)
        
        If lstA.Selected(i) Then
            If lnI = 0 Then
                lnI = lnI + 1
                If Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "datetime" Then
                    sqlAux = "Convert(Varchar(10)," & Trim(Left(Me.lstA.List(i), 150)) & ",103)" & Trim(Left(Me.lstA.List(i), 150))
                Else
                    sqlAux = Trim(Left(Me.lstA.List(i), 150))
                End If
            Else
                If Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "datetime" Then
                    sqlAux = sqlAux & " , Convert(Varchar(10)," & Trim(Left(Me.lstA.List(i), 150)) & ",103)" & Trim(Left(Me.lstA.List(i), 150))
                Else
                    sqlAux = sqlAux & " , " & Trim(Left(Me.lstA.List(i), 150))
                End If
            End If
        End If
        
        If Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "datetime" Then
            If lnPos <> 0 Then
                If lbBanFiltro Then
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                            sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Between '" & Format(CDate(Mid(Me.lstA.List(i), lnPos + 1, 10)), gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, CDate(Mid(Me.lstA.List(i), lnPosOtro + 1, 10))), gsFormatoFecha) & "' "
                        Else
                            sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Not Between '" & Format(CDate(Mid(Me.lstA.List(i), lnPos + 1, 10)), gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, CDate(Mid(Me.lstA.List(i), lnPosOtro + 1, 10))), gsFormatoFecha) & "' "
                        End If
                    Else
                        sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Between '" & Format(CDate(Mid(Me.lstA.List(i), lnPos + 1, 10)), gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, CDate(Mid(Me.lstA.List(i), lnPosOtro + 1, 10))), gsFormatoFecha) & "' "
                    End If
                    lbBanFiltro = False
                Else
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Between '" & Format(CDate(Mid(Me.lstA.List(i), lnPos + 1, 10)), gsFormatoFecha) & "' And '" & Format(CDate(Mid(Me.lstA.List(i), lnPosOtro + 1, 10)), gsFormatoFecha) & "' "
                        Else
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Not Between '" & Format(CDate(Mid(Me.lstA.List(i), lnPos + 1, 10)), gsFormatoFecha) & "' And '" & Format(CDate(Mid(Me.lstA.List(i), lnPosOtro + 1, 10)), gsFormatoFecha) & "' "
                        End If
                    End If
                End If
            End If
        ElseIf Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "char" Or Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "varchar" Then
            If lnPos <> 0 Then
                If lbBanFiltro Then
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPos + 1, 1) = "0" Then
                            If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                                sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "' "
                            Else
                                sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Not Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "' "
                            End If
                        Else
                            If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                                sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "'"
                            Else
                                sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Not Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "'"
                            End If
                        End If
                    End If
                    lbBanFiltro = False
                Else
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPos + 1, 1) = "0" Then
                            If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                                sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                            Else
                                sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Not Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                            End If
                        Else
                            If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                                sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                            Else
                                sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Not Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                            End If
                        End If
                    Else
                        If Mid(Me.lstA.List(i), lnPos + 1, 1) = "0" Then
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                        Else
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Like '" & Mid(Me.lstA.List(i), lnPos + 2, lnPosDiff - lnPos - 2) & "%' "
                        End If
                    End If
                End If
            End If
        ElseIf Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "numeric" Or Mid(Me.lstA.List(i), lnPosIni, lnPosFin - lnPosIni) = "money" Then
            If lnPos <> 0 Then
                If lbBanFiltro Then
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                            sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Between " & Format(lnPosOtro - lnPos - 1, "#.00") & " And " & Format(Mid(Me.lstA.List(i), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1), "#.00") & " "
                        Else
                            sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Not Between " & Format(lnPosOtro - lnPos - 1, "#.00") & " And " & Format(Mid(Me.lstA.List(i), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1), "#.00") & " "
                        End If
                    Else
                        sqlFilto = " Where " & Trim(Left(Me.lstA.List(i), 150)) & " Between " & Format(lnPosOtro - lnPos - 1, "#.00") & " And " & Format(Mid(Me.lstA.List(i), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1), "#.00") & " "
                    End If
                    lbBanFiltro = False
                Else
                    If lnPosDiff <> 0 Then
                        If Mid(Me.lstA.List(i), lnPosDiff + 1, 1) = "=" Then
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Between " & Format(lnPosOtro - lnPos - 1, "#.00") & " And " & Format(Mid(Me.lstA.List(i), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1), "#.00") & " "
                        Else
                            sqlFilto = sqlFilto & " And " & Trim(Left(Me.lstA.List(i), 150)) & " Not Between " & Format(lnPosOtro - lnPos - 1, "#.00") & " And " & Format(Mid(Me.lstA.List(i), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1), "#.00") & " "
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    sqlP = "Select Distinct " & sqlAux & " From ReporteGeneral " & sqlFilto
    Set rsP = oRep.GetRepGen(sqlP)
    
    lsArchivoN = App.path & "\Spooler\" & gsCodUser & Format(CDate(gdFecSis), "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
        
        i = 0
        While Not rsP.EOF
            i = i + 1
            If i = 1 Then
                For J = 0 To rsP.Fields.Count - 1
                    xlHoja1.Cells(i, J + 1) = rsP.Fields(J).Name
                Next J
            Else
                For J = 0 To rsP.Fields.Count - 1
                    If rsP.Fields(J).Type = adDBTimeStamp Then
                        xlHoja1.Cells(i, J + 1) = Format(rsP.Fields(J), gsFormatoFechaView)
                    Else
                        xlHoja1.Cells(i, J + 1) = rsP.Fields(J)
                    End If
                Next J
                rsP.MoveNext
            End If
        Wend
    
        xlHoja1.Range("1:1").Font.Bold = True
         
        xlHoja1.Columns.AutoFit
         
        OleExcel.Class = "ExcelWorkSheet"
        ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
        OleExcel.SourceDoc = lsArchivoN
        OleExcel.Verb = 1
        OleExcel.Action = 1
        OleExcel.DoVerb -1
        
        
    End If
    MousePointer = 0
        
    rsP.Close
    Set rsP = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsR As ADODB.Recordset
    Set rsR = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    
    Set rsR = oRep.GetReporteGralCampos
    
    Me.lstA.Clear
    Me.lstP.Clear
    
    If Not (rsR.EOF And rsR.BOF) Then
        While Not rsR.EOF
            lstP.AddItem rsR!COLUMN_NAME & Space(150) & "," & rsR!TYPE_NAME & "-" & rsR!DATA_TYPE
            rsR.MoveNext
        Wend
    End If
    
    rsR.Close
    Set rsR = Nothing
    
    Me.fraFechas.Visible = False
    Me.fraNumeros.Visible = False
    Me.fratexto.Visible = False
    Me.optLista.value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub lstA_Click()
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnPos  As Integer
    
    Dim lnPosDiff As Integer
    Dim lnPosOtro As Integer
    
    
    If lstA.ListIndex = -1 Then Exit Sub
    
    lnPos = InStr(1, Me.lstA.List(lstA.ListIndex), "*", vbTextCompare)
    lnPosIni = InStr(1, Me.lstA.List(lstA.ListIndex), ",", vbTextCompare) + 1
    lnPosFin = InStr(1, Me.lstA.List(lstA.ListIndex), "-", vbTextCompare)
    lnPosDiff = InStr(1, Me.lstA.List(lstA.ListIndex), "{", vbTextCompare)
    lnPosOtro = InStr(1, Me.lstA.List(lstA.ListIndex), "$", vbTextCompare)
    
    
    If Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "datetime" Then
        If lnPos = 0 Then
            Me.mskIni.Text = "__/__/____"
            Me.mskFin.Text = "__/__/____"
        Else
            Me.mskIni.Text = Mid(Me.lstA.List(lstA.ListIndex), lnPos + 1, lnPosOtro - lnPos - 1)
            Me.mskFin.Text = Mid(Me.lstA.List(lstA.ListIndex), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1)
        End If
        Me.fraFechas.Visible = True
        Me.fraNumeros.Visible = False
        Me.fratexto.Visible = False
    ElseIf Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "numeric" Or Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "money" Then
        If lnPos = 0 Then
            Me.txtDesde.Text = ""
            Me.txtHasta.Text = ""
        Else
            Me.txtDesde.Text = Mid(Me.lstA.List(lstA.ListIndex), lnPos + 1, lnPosOtro - lnPos - 1)
            Me.txtHasta.Text = Mid(Me.lstA.List(lstA.ListIndex), lnPosOtro + 1, lnPosDiff - lnPosOtro - 1)
        End If
        Me.fraFechas.Visible = False
        Me.fraNumeros.Visible = True
        Me.fratexto.Visible = False
    ElseIf Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "char" Or Mid(Me.lstA.List(lstA.ListIndex), lnPosIni, lnPosFin - lnPosIni) = "varchar" Then
        Me.fraFechas.Visible = False
        Me.fraNumeros.Visible = False
        Me.fratexto.Visible = True
        LLenaLista Left(lstA.List(lstA.ListIndex), 150)
        Me.txtTexto.Text = ""
        If lnPos = 0 Then
            Me.cmbLista.ListIndex = -1
        Else
            If Mid(Me.lstA.List(lstA.ListIndex), lnPos + 1, 1) = "0" Then
                Me.optTexto.value = True
                Me.txtTexto.Text = Mid(Me.lstA.List(lstA.ListIndex), lnPos + 2, lnPosDiff - lnPos - 3)
            Else
                Me.optLista.value = True
                UbicaCombo cmbLista, Mid(Me.lstA.List(lstA.ListIndex), lnPos + 2, lnPosDiff - lnPos - 2)
            End If
        End If
    Else
        Me.fraFechas.Visible = False
        Me.fraNumeros.Visible = True
        Me.fratexto.Visible = False
    End If
    If lnPos <> 0 Then
        If Mid(Me.lstA.List(lstA.ListIndex), lnPosDiff + 1, 1) = "=" Then
            Me.optDiff = False
            Me.optIgual = True
        Else
            Me.optDiff = True
            Me.optIgual = False
        End If
    Else
            Me.optDiff = False
            Me.optIgual = True
    End If
End Sub

Private Sub lstA_DblClick()
    cmdIzqU_Click
End Sub

Private Sub lstP_DblClick()
    cmdDerU_Click
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAplicar.SetFocus
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFin.SetFocus
    End If
End Sub

Private Sub optLista_Click()
    If optLista.value Then
        Me.txtTexto.Visible = False
        Me.cmbLista.Visible = True
    Else
        Me.txtTexto.Visible = True
        Me.cmbLista.Visible = False
        Me.cmbLista.ListIndex = -1
    End If
End Sub

Private Sub optTexto_Click()
    If optLista.value Then
        Me.txtTexto.Visible = False
        Me.cmbLista.Visible = True
    Else
        Me.txtTexto.Visible = True
        Me.cmbLista.Visible = False
        Me.cmbLista.ListIndex = -1
    End If
End Sub

Private Sub LLenaLista(psCodigo As String)
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsC = oRep.GetReporteGralValores(psCodigo)
    Me.cmbLista.Clear
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            If Not IsNull(rsC.Fields(0)) Then Me.cmbLista.AddItem rsC.Fields(0)
            rsC.MoveNext
        Wend
    End If
    Me.cmbLista.AddItem ""
    rsC.Close
    Set rsC = Nothing
End Sub

Public Sub Ini(psCaption As String, pForm As Form)
    Caption = psCaption
    Show 0, pForm
End Sub

Private Sub txtDesde_GotFocus()
    txtDesde.SelStart = 0
    txtDesde.SelLength = 50
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtHasta.SetFocus
    End If
End Sub

Private Sub txtHasta_GotFocus()
    txtHasta.SelStart = 0
    txtHasta.SelLength = 50
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAplicar.SetFocus
    End If
End Sub
