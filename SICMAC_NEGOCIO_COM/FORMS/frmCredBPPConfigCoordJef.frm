VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPConfigCoordJef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuración Coordinadores de Créditos, Jefes de Agencia y Jefes Territoriales"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "frmCredBPPConfigCoordJef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Bonificación:"
      TabPicture(0)   =   "frmCredBPPConfigCoordJef.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGeneral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardarCG"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdGuardarCG 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   9960
         TabIndex        =   10
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Coordinadores y Jefes de Agencia"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   3030
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Cargo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   5
               Top             =   315
               Width           =   495
            End
         End
         Begin SICMACT.FlexEdit feParemetros 
            Height          =   1575
            Left            =   120
            TabIndex        =   6
            Top             =   765
            Width           =   10680
            _extentx        =   18838
            _extenty        =   2778
            cols0           =   9
            highlight       =   1
            encabezadosnombres=   "#-Cargo-Desde-Hasta-Desde-Hasta-Minimo(%)-Maximo(%)-Aux"
            encabezadosanchos=   "0-3000-1200-1200-1200-1200-1200-1200-0"
            font            =   "frmCredBPPConfigCoordJef.frx":0326
            font            =   "frmCredBPPConfigCoordJef.frx":034E
            font            =   "frmCredBPPConfigCoordJef.frx":0376
            font            =   "frmCredBPPConfigCoordJef.frx":039E
            font            =   "frmCredBPPConfigCoordJef.frx":03C6
            fontfixed       =   "frmCredBPPConfigCoordJef.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-2-3-4-5-6-7-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-R-R-R-L"
            formatosedit    =   "0-0-3-3-3-3-2-2-0"
            cantentero      =   6
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "% de Bonificación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7935
            TabIndex        =   9
            Top             =   450
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Maximo de Analistas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5535
            TabIndex        =   8
            Top             =   450
            Width           =   2415
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Minimo de Analistas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3135
            TabIndex        =   7
            Top             =   450
            Width           =   2415
         End
      End
      Begin VB.Frame fraGeneral 
         Caption         =   "Jefes Territoriales"
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
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   10935
         Begin SICMACT.EditMoney txtPorBonifiacionJN 
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   360
            Width           =   1095
            _extentx        =   1931
            _extenty        =   450
            font            =   "frmCredBPPConfigCoordJef.frx":0414
            text            =   "0"
            enabled         =   -1
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Minimo a Bonificar de sus Jefes de Agencia(%):"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3330
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPConfigCoordJef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'Private j As Integer
'
'Private Sub CargaControles()
'Dim oPar As COMDCredito.DCOMParametro
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'Dim rsDBPPDatos As ADODB.Recordset
'Dim oDBPP As COMDCredito.DCOMBPPR
'Set oDBPP = New COMDCredito.DCOMBPPR
'Set oConst = New COMDConstantes.DCOMConstantes
'
''Parametros
'LimpiaFlex feParemetros
'Set rsConst = oConst.RecuperaConstantes(7072)
'If Not (rsConst.EOF And rsConst.BOF) Then
'    For i = 0 To rsConst.RecordCount - 1
'        feParemetros.AdicionaFila
'        feParemetros.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'        feParemetros.TextMatrix(i + 1, 8) = Trim(rsConst!nConsValor)
'        Set rsDBPPDatos = oDBPP.ObtenerParametrosCoorgJefAg(CInt(Trim(rsConst!nConsValor)))
'        If Not (rsDBPPDatos.EOF And rsDBPPDatos.BOF) Then
'            feParemetros.TextMatrix(i + 1, 2) = Trim(rsDBPPDatos!nDesdeMinA)
'            feParemetros.TextMatrix(i + 1, 3) = Trim(rsDBPPDatos!nHastaMinA)
'            feParemetros.TextMatrix(i + 1, 4) = Trim(rsDBPPDatos!nDesdeMaxA)
'            feParemetros.TextMatrix(i + 1, 5) = Trim(rsDBPPDatos!nHastaMaxA)
'            feParemetros.TextMatrix(i + 1, 6) = Format(Trim(rsDBPPDatos!nMinBonificacion), "#00.00")
'            feParemetros.TextMatrix(i + 1, 7) = Format(Trim(rsDBPPDatos!nMaxBonificacion), "#00.00")
'            Set rsDBPPDatos = Nothing
'        End If
'        rsConst.MoveNext
'    Next i
'    Set rsConst = Nothing
'End If
'
'Set oPar = New COMDCredito.DCOMParametro
'txtPorBonifiacionJN.Text = Format(oPar.RecuperaValorParametro(3212), "#00.00")
'Set oPar = Nothing
'
'End Sub
'
'Private Sub cmdGuardarCG_Click()
'On Error GoTo Error
'If ValidaDatos Then
'    If MsgBox("Estas seguro de Guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMDCredito.DCOMBPPR
'        Dim oPar As COMDCredito.DCOMParametro
'
'        Set oPar = New COMDCredito.DCOMParametro
'        Set oBPP = New COMDCredito.DCOMBPPR
'
'        Call oBPP.OpeParametrosCoorgJefAg(2)
'
'        For i = 0 To feParemetros.Rows - 2
'            Call oBPP.OpeParametrosCoorgJefAg(1, CInt(feParemetros.TextMatrix(i + 1, 8)), CLng(feParemetros.TextMatrix(i + 1, 2)), _
'            CLng(feParemetros.TextMatrix(i + 1, 3)), CLng(feParemetros.TextMatrix(i + 1, 4)), CLng(feParemetros.TextMatrix(i + 1, 5)), _
'            CDbl(feParemetros.TextMatrix(i + 1, 6)), CDbl(feParemetros.TextMatrix(i + 1, 7)))
'        Next i
'
'        Call oPar.ModificarParametro("3212", "% Min a Bonificar de Jef. Ag.(Jef. Territoriales)", CDbl(txtPorBonifiacionJN.Text), "")
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'        CargaControles
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub Form_Load()
'CargaControles
'End Sub
'
'Private Function ValidaDatos() As Boolean
'ValidaDatos = True
'Dim lsDato As String
'
'If Trim(feParemetros.TextMatrix(1, 0)) <> "" Then
'    For i = 0 To feParemetros.Rows - 2
'        For j = 2 To 7
'            If Trim(feParemetros.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Favor de Ingresar todos los campos Correctamente en los parámetros de ''" & feParemetros.TextMatrix(i + 1, 1) & "''", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If CDbl(Trim(feParemetros.TextMatrix(i + 1, j))) > 99999999 Then
'                Select Case j
'                    Case 2: lsDato = "Valor Inicial del Nº Minimo de Analistas"
'                    Case 3: lsDato = "Valor Final del Nº Minimo de Analistas"
'                    Case 4: lsDato = "Valor Inicial del Nº Maximo de Analistas"
'                    Case 5: lsDato = "Valor Final del Nº Maximo de Analistas"
'                    Case 6: lsDato = "Valor del % Minimo de la Bonificación"
'                    Case 7: lsDato = "Valor del % Maximo de la Bonificación"
'                End Select
'                MsgBox "El " & lsDato & "  Supera el Valor Máximo Permitido", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Then
'                If IsNumeric(feParemetros.TextMatrix(i + 1, j)) Then
'                    If CLng(Trim(feParemetros.TextMatrix(i + 1, j))) < 1 Then
'                        Select Case j
'                            Case 3: lsDato = "Valor Final del Nº Minimo de Analistas"
'                            Case 4: lsDato = "Valor Inicial del Nº Maximo de Analistas"
'                            Case 5: lsDato = "Valor Final del Nº Maximo de Analistas"
'                            Case 6: lsDato = "Valor del % Minimo de la Bonificación"
'                            Case 7: lsDato = "Valor del % Maximo de la Bonificación"
'                        End Select
'                        MsgBox "El " & lsDato & " en los parámetros de ''" & feParemetros.TextMatrix(i + 1, 1) & "'' debe ser mayor a 0.", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            End If
'
'            If j = 6 Or j = 7 Then
'                 If IsNumeric(feParemetros.TextMatrix(i + 1, j)) Then
'                    If CLng(Trim(feParemetros.TextMatrix(i + 1, j))) > 100 Then
'                        Select Case j
'                            Case 6: lsDato = "Valor del % Minimo de la Bonificación"
'                            Case 7: lsDato = "Valor del % Maximo de la Bonificación"
'                        End Select
'                        MsgBox "El " & lsDato & " en los parámetros de ''" & feParemetros.TextMatrix(i + 1, 1) & "'' no debe ser mayor a 100.", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next j
'
'        If IsNumeric(feParemetros.TextMatrix(i + 1, 3)) And IsNumeric(feParemetros.TextMatrix(i + 1, 4)) Then
'            If (CLng(Trim(feParemetros.TextMatrix(i + 1, 4))) - CLng(Trim(feParemetros.TextMatrix(i + 1, 3)))) <> 1 Then
'
'                MsgBox "La diferencia entre el Valor final del Nº Minimo de Analista y el Valor Inicial" & Chr(10) & _
'                        "del Nº Maximo de Analistas en los parámetros de ''" & feParemetros.TextMatrix(i + 1, 1) & "'' deber ser 1.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'    Next i
'End If
'
'If IsNumeric(txtPorBonifiacionJN.Text) Then
'    If CDbl(txtPorBonifiacionJN.Text) = 0 Then
'        MsgBox "Favor de Ingresar el Valor del Minimo a Bonificar para los Jefes Territoriales.", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'End If
'
'End Function
'
'Private Sub txtPorBonifiacionJN_Change()
'   If Trim(txtPorBonifiacionJN.Text) <> "." Then
'        If CDbl(txtPorBonifiacionJN.Text) > 100 Then
'            txtPorBonifiacionJN.Text = "100.00"
'        End If
'
'         If CDbl(txtPorBonifiacionJN.Text) < 0 Then
'            txtPorBonifiacionJN.Text = "0.00"
'        End If
'    Else
'        txtPorBonifiacionJN.Text = "0.00"
'    End If
'End Sub
'
'Private Sub txtPorBonifiacionJN_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdGuardarCG.SetFocus
'End If
'End Sub
