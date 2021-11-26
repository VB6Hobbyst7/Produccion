VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPBonoConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Consulta de Bonos Cerrados"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "frmCredBPPBonoConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   2250
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3969
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consultar Bonos"
      TabPicture(0)   =   "frmCredBPPBonoConsulta.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBusqueda"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCerrar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
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
         Left            =   4680
         TabIndex        =   4
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Filtro"
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
         TabIndex        =   1
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton cmdConsultar 
            Caption         =   "Consultar"
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
            Left            =   4560
            TabIndex        =   5
            Top             =   600
            Width           =   1170
         End
         Begin VB.ComboBox cmbMeses 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   600
            Width           =   1935
         End
         Begin SICMACT.uSpinner spAnioBP 
            Height          =   315
            Left            =   3120
            TabIndex        =   6
            Top             =   600
            Width           =   1215
            _extentx        =   2143
            _extenty        =   556
            max             =   9999
            min             =   1900
            maxlength       =   4
            min             =   1900
            font            =   "frmCredBPPBonoConsulta.frx":0326
            fontname        =   "MS Sans Serif"
            fontsize        =   8.25
         End
         Begin VB.Label lblFiltroSelect 
            AutoSize        =   -1  'True
            Caption         =   "Mes-Año:"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPBonoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdConsultar_Click()
'If Trim(cmbMeses.Text) = "" Then
'    MsgBox "Seleccione el mes", vbInformation, "Aviso"
'    cmbMeses.SetFocus
'    Exit Sub
'End If
'
'If Trim(spAnioBP.valor) = "" Or spAnioBP.valor = 0 Then
'    MsgBox "Ingrese el Año", vbInformation, "Aviso"
'    spAnioBP.SetFocus
'    Exit Sub
'End If
'
'
'Dim fgFecActual As Date
'fgFecActual = CDate(Trim(Right(cmbMeses.Text, 15)))
'Call frmCredBPPBonoCierres.Inicio(Month(fgFecActual), Year(fgFecActual), "BPP - Consulta de Bonos Generados (Al Cierre " & Format(fgFecActual, "DD/MM/YYYY") & ")", False)
'End Sub
'
'Private Sub CargaCombo()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'
'Set rsBPP = oBPP.ObtenerCierreBonoGenerados(CInt(spAnioBP.valor))
'
'cmbMeses.Clear
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        cmbMeses.AddItem MesAnio(CDate(rsBPP!dFechaCierre)) & Space(75) & Format(CDate(rsBPP!dFechaCierre), "DD/MM/YYYY")
'        rsBPP.MoveNext
'    Next i
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'End Sub
'
'Private Sub Form_Load()
'spAnioBP.valor = Year(gdFecSis)
'CargaCombo
'End Sub
'
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
'sFechaDesc = sFechaDesc
'MesAnio = UCase(sFechaDesc)
'End Function
'
'Private Sub spAnioBP_Change()
'CargaCombo
'End Sub
