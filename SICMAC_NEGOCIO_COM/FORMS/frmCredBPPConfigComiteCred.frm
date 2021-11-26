VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPConfigComiteCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuración de Comités de Crédito"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "frmCredBPPConfigComiteCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Coordinadores y Analistas"
      TabPicture(0)   =   "frmCredBPPConfigComiteCred.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraNuevaMora"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fraNuevaMora 
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
         ForeColor       =   &H8000000D&
         Height          =   3855
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   8775
         Begin SICMACT.FlexEdit feComite 
            Height          =   3495
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   8520
            _extentx        =   15028
            _extenty        =   6165
            cols0           =   8
            highlight       =   1
            encabezadosnombres=   "#-PersCod-Act.-Usuario-Nombre-Comité-Cargos-Aux"
            encabezadosanchos=   "0-0-500-1000-3500-1500-1500-0"
            font            =   "frmCredBPPConfigComiteCred.frx":0326
            font            =   "frmCredBPPConfigComiteCred.frx":034E
            font            =   "frmCredBPPConfigComiteCred.frx":0376
            font            =   "frmCredBPPConfigComiteCred.frx":039E
            font            =   "frmCredBPPConfigComiteCred.frx":03C6
            fontfixed       =   "frmCredBPPConfigComiteCred.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-2-X-X-5-6-X"
            listacontroles  =   "0-0-4-0-0-3-3-0"
            encabezadosalineacion=   "C-C-C-C-L-L-L-C"
            formatosedit    =   "0-0-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
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
         Left            =   7800
         TabIndex        =   2
         Top             =   4800
         Width           =   1170
      End
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
         Left            =   6480
         TabIndex        =   1
         Top             =   4800
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes a generar el BPP:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1590
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         Caption         =   "@Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCredBPPConfigComiteCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fgFecActual As Date
'Private i As Integer
'
'Private Sub CargaControles()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.ObtenerComiteAnaCord(gsCodAge, Month(fgFecActual), Year(fgFecActual))
'
'LimpiaFlex feComite
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 0 To rsBPP.RecordCount - 1
'        feComite.AdicionaFila
'        feComite.TextMatrix(i + 1, 0) = i + 1
'        feComite.TextMatrix(i + 1, 1) = Trim(rsBPP!cPersCod)
'        feComite.TextMatrix(i + 1, 2) = Trim(rsBPP!nEstado)
'        feComite.TextMatrix(i + 1, 3) = Trim(rsBPP!cUser)
'        feComite.TextMatrix(i + 1, 4) = Trim(rsBPP!cPersNombre)
'        feComite.TextMatrix(i + 1, 5) = Trim(rsBPP!comite) & Space(75) & Trim(rsBPP!CodComite)
'        feComite.TextMatrix(i + 1, 6) = Trim(rsBPP!Cargo) & Space(75) & Trim(rsBPP!CodCargo)
'        feComite.TextMatrix(i + 1, 7) = ""
'        rsBPP.MoveNext
'    Next i
'Else
'    MsgBox "No Hay Datos.", vbInformation, "Aviso"
'End If
'feComite.TopRow = 1
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'If MsgBox("Estas seguro de Guardar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim i As Integer
'    Dim nEstado As Integer
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    For i = 0 To feComite.Rows - 2
'        nEstado = CInt(IIf(Trim(feComite.TextMatrix(i + 1, 2)) = ".", 1, 0))
'        Call oBPP.OpeComiteAnaCord(feComite.TextMatrix(i + 1, 1), CInt(Trim(Right(feComite.TextMatrix(i + 1, 5), 4))), gsCodAge, Month(fgFecActual), Year(fgFecActual), gsCodUser, gdFecSis, nEstado, CInt(Trim(Right(feComite.TextMatrix(i + 1, 6), 4))))
'    Next i
'    Call CargaControles
'    MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'End If
'End Sub
'
'
'Private Sub feComite_RowColChange()
'Dim oConst As COMDConstantes.DCOMConstantes
'Set oConst = New COMDConstantes.DCOMConstantes
'
'Select Case feComite.Col
'    Case 5: feComite.CargaCombo oConst.RecuperaConstantes(7066)
'    Case 6: feComite.CargaCombo oConst.RecuperaConstantes(7074)
'End Select
'
'Set oConst = Nothing
'End Sub
'
'Private Sub Form_Load()
'MesActual
'CargaControles
'lblMes.Caption = MesAnio(fgFecActual)
'End Sub
'Private Sub MesActual()
'    Dim oConsSist As COMDConstSistema.NCOMConstSistema
'    Set oConsSist = New COMDConstSistema.NCOMConstSistema
'    fgFecActual = oConsSist.LeeConstSistema(gConstSistFechaBPP)
'    Set oConsSist = Nothing
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
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
