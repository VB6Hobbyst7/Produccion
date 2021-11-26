VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPMoraBaseCamb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Cambiar Mora Base"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "frmCredBPPMoraBaseCamb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNuevaMora 
      Caption         =   "Nueva Mora"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
      Begin VB.ComboBox cmbMesAnioNew 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblMoraBaseNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblMesMoraNew 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblMoraBaseNewDesc 
         AutoSize        =   -1  'True
         Caption         =   "Mora Base:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   810
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Moras"
      TabPicture(0)   =   "frmCredBPPMoraBaseCamb.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCerrar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraMoraActual"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Frame fraMoraActual 
         Caption         =   "Mora Actual"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         Begin VB.Label lblMoraBase 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   10
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblMesMoraBaseAct 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblMoraBaseAct 
            AutoSize        =   -1  'True
            Caption         =   "Mora Base:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   810
         End
         Begin VB.Label lblMesMoraAct 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   3120
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCredBPPMoraBaseCamb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fpMatMoraBase As MoraBase
'
'Private fdFecMoraBase As Date
'Private fnPorcMoraBase As Double
'
'Public Sub Inicio(ByRef pMatMoraBase As MoraBase)
'fpMatMoraBase = pMatMoraBase
'
'lblMesMoraBaseAct.Caption = MesAnio(pMatMoraBase.FecMoraBase)
'lblMoraBase.Caption = CStr(Format(Round(pMatMoraBase.PorcMoraBase * 100, 2), "##0.00")) & " % "
'
'Call CargaCombo(fpMatMoraBase)
'Me.Show 1
'End Sub
'
'Private Sub cmbMesAnioNew_Click()
'lblMoraBaseNew.Caption = ObtenerValor(Trim(Right(cmbMesAnioNew.Text, 20)))
'End Sub
'Private Function ObtenerValor(ByVal pFecha As Date)
'Dim sValor As String
'
'Select Case pFecha
'    Case fpMatMoraBase.FechaEne:
'        fdFecMoraBase = fpMatMoraBase.FechaEne
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseEne
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseEne * 100, 2), "##0.00")) & "% "
'   Case fpMatMoraBase.FechaFeb:
'        fdFecMoraBase = fpMatMoraBase.FechaFeb
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseFeb
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseFeb * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaMar:
'        fdFecMoraBase = fpMatMoraBase.FechaMar
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseMar
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseMar * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaAbr:
'        fdFecMoraBase = fpMatMoraBase.FechaAbr
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseAbr
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseAbr * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaMay:
'        fdFecMoraBase = fpMatMoraBase.FechaMay
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseMay
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseMay * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaJun:
'        fdFecMoraBase = fpMatMoraBase.FechaJun
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseJun
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseJun * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaJul:
'        fdFecMoraBase = fpMatMoraBase.FechaJul
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseJul
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseJul * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaAgo:
'        fdFecMoraBase = fpMatMoraBase.FechaAgo
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseAgo
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseAgo * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaSep:
'        fdFecMoraBase = fpMatMoraBase.FechaSep
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseSep
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseSep * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaOct:
'        fdFecMoraBase = fpMatMoraBase.FechaOct
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseOct
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseOct * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaNov:
'        fdFecMoraBase = fpMatMoraBase.FechaNov
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseNov
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseNov * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaDic:
'        fdFecMoraBase = fpMatMoraBase.FechaDic
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseDic
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseDic * 100, 2), "##0.00")) & "% "
'    Case fpMatMoraBase.FechaDicAnt:
'        fdFecMoraBase = fpMatMoraBase.FechaDicAnt
'        fnPorcMoraBase = fpMatMoraBase.PorcMoraBaseDicAnt
'        sValor = CStr(Format(Round(fpMatMoraBase.PorcMoraBaseDicAnt * 100, 2), "##0.00")) & "% "
'End Select
'
'ObtenerValor = sValor
'End Function
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'Private Function MesAnio(ByVal dFecha As Date) As String
'Dim sFechaDesc As String
'sFechaDesc = ""
'
'Select Case Month(dFecha)
'    Case 1: sFechaDesc = "Enero"
'    Case 2: sFechaDesc = "Febrero"
'    Case 3: sFechaDesc = "Marzo"
'    Case 4: sFechaDesc = "Abril"
'    Case 5: sFechaDesc = "Mayo"
'    Case 6: sFechaDesc = "Junio"
'    Case 7: sFechaDesc = "Julio"
'    Case 8: sFechaDesc = "Agosto"
'    Case 9: sFechaDesc = "Septiembre"
'    Case 10: sFechaDesc = "Octubre"
'    Case 11: sFechaDesc = "Noviembre"
'    Case 12: sFechaDesc = "Diciembre"
'End Select
'
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
'Private Sub CargaCombo(ByRef pMatMoraBase As MoraBase)
'cmbMesAnioNew.Clear
'
'    If pMatMoraBase.FechaDicAnt <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaDicAnt) & Space(75) & Trim((pMatMoraBase.FechaDicAnt))
'    End If
'    If pMatMoraBase.FechaEne <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaEne) & Space(75) & Trim((pMatMoraBase.FechaEne))
'    End If
'    If pMatMoraBase.FechaFeb <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaFeb) & Space(75) & Trim((pMatMoraBase.FechaFeb))
'    End If
'    If pMatMoraBase.FechaMar <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaMar) & Space(75) & Trim((pMatMoraBase.FechaMar))
'    End If
'    If pMatMoraBase.FechaAbr <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaAbr) & Space(75) & Trim((pMatMoraBase.FechaAbr))
'    End If
'    If pMatMoraBase.FechaMay <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaMay) & Space(75) & Trim((pMatMoraBase.FechaMay))
'    End If
'    If pMatMoraBase.FechaJun <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaJun) & Space(75) & Trim((pMatMoraBase.FechaJun))
'    End If
'    If pMatMoraBase.FechaJul <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaJul) & Space(75) & Trim((pMatMoraBase.FechaJul))
'    End If
'    If pMatMoraBase.FechaAgo <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaAgo) & Space(75) & Trim((pMatMoraBase.FechaAgo))
'    End If
'    If pMatMoraBase.FechaSep <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaSep) & Space(75) & Trim((pMatMoraBase.FechaSep))
'    End If
'    If pMatMoraBase.FechaOct <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaOct) & Space(75) & Trim((pMatMoraBase.FechaOct))
'    End If
'    If pMatMoraBase.FechaNov <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaNov) & Space(75) & Trim((pMatMoraBase.FechaNov))
'    End If
'    If pMatMoraBase.FechaDic <> "01/01/1900" Then
'        cmbMesAnioNew.AddItem MesAnio(pMatMoraBase.FechaDic) & Space(75) & Trim((pMatMoraBase.FechaDic))
'    End If
'End Sub
'
'Private Sub cmdGuardar_Click()
'If Trim(cmbMesAnioNew.Text) = "" Then
'    MsgBox "Selecciona una Nueva Mora Base", vbInformation, "Aviso"
'    cmbMesAnioNew.SetFocus
'    Exit Sub
'End If
'
'Dim oBPP As COMDCredito.DCOMBPPR
'Set oBPP = New COMDCredito.DCOMBPPR
'
'If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Call oBPP.OpeMoraBase(fpMatMoraBase.codigo, fdFecMoraBase, fnPorcMoraBase, Year(fpMatMoraBase.Fecha), Month(fpMatMoraBase.Fecha), 1)
'    MsgBox "Se guardaron los cambios correctamente", vbInformation, "Aviso"
'    frmCredBPPMoraBase.LimpiaDatos
'    Unload Me
'End If
'End Sub
