VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPMoraBaseDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Detalle de Mora Base"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmCredBPPMoraBaseDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Moras"
      TabPicture(0)   =   "frmCredBPPMoraBaseDet.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMesMora"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMesMoraBase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feMoraDetalle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
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
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feMoraDetalle 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5760
         _extentx        =   10160
         _extenty        =   2778
         cols0           =   6
         highlight       =   1
         encabezadosnombres=   "#-Cartera-Aux-Saldo-Mora-%"
         encabezadosanchos=   "0-1000-0-2200-1200-1000"
         font            =   "frmCredBPPMoraBaseDet.frx":0326
         font            =   "frmCredBPPMoraBaseDet.frx":034E
         font            =   "frmCredBPPMoraBaseDet.frx":0376
         font            =   "frmCredBPPMoraBaseDet.frx":039E
         font            =   "frmCredBPPMoraBaseDet.frx":03C6
         fontfixed       =   "frmCredBPPMoraBaseDet.frx":03EE
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-C-C"
         formatosedit    =   "0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
      End
      Begin VB.Label lblMesMoraBase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   510
         Width           =   2055
      End
      Begin VB.Label lblMesMora 
         AutoSize        =   -1  'True
         Caption         =   "Mes Mora Base:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   570
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmCredBPPMoraBaseDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim i As Integer
'Private fpMatMoraBase As MoraBase
'Public Sub Inicio(ByRef pMatMoraBase As MoraBase)
'CargaControles
'fpMatMoraBase = pMatMoraBase
'
'lblMesMoraBase.Caption = MesAnio(fpMatMoraBase.FecMoraBase)
'
'Call MostrarDatos
'Me.Show 1
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub CargaControles()
'Call LimpiaFlex(feMoraDetalle)
'
'    For i = 1 To 4
'        feMoraDetalle.AdicionaFila
'        feMoraDetalle.TextMatrix(i, 1) = Choose(i, "Propia", "Heredado", "Salida", "Final")
'        feMoraDetalle.TextMatrix(i, 2) = Choose(i, "1", "2", "3", "4")
'     Next i
'End Sub
'
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
'MesAnio = sFechaDesc
'End Function
'
'Private Sub MostrarDatos()
'Dim oBPP As COMDCredito.DCOMBPPR
'Dim rsBPP As ADODB.Recordset
'Dim Resultado As Double
'
'Set oBPP = New COMDCredito.DCOMBPPR
'
'Call LimpiaFlex(feMoraDetalle)
'
'CargaControles
'
'For i = 1 To 4
'    feMoraDetalle.TextMatrix(i, 3) = Format(Choose(i, Round(fpMatMoraBase.SaldoPro, 2), Round(fpMatMoraBase.SaldoHer, 2), Round(fpMatMoraBase.SaldoSal, 2), Round(fpMatMoraBase.Saldo, 2)), "###," & String(15, "#") & "#0.00")
'    feMoraDetalle.TextMatrix(i, 4) = Format(Choose(i, Round(fpMatMoraBase.MoraPro, 2), Round(fpMatMoraBase.MoraHer, 2), Round(fpMatMoraBase.MoraSal, 2), Round(fpMatMoraBase.Mora, 2)), "###," & String(15, "#") & "#0.00")
'
'    Resultado = 0
'    Select Case i
'        Case 1:
'                If fpMatMoraBase.SaldoPro = 0 Then
'                    Resultado = 0
'                Else
'                    Resultado = Round(fpMatMoraBase.MoraPro / fpMatMoraBase.SaldoPro * 100, 2)
'                End If
'        Case 2:
'                If fpMatMoraBase.SaldoHer = 0 Then
'                    Resultado = 0
'                Else
'                    Resultado = Round(fpMatMoraBase.MoraHer / fpMatMoraBase.SaldoHer * 100, 2)
'                End If
'        Case 3:
'                If fpMatMoraBase.SaldoSal = 0 Then
'                    Resultado = 0
'                Else
'                    Resultado = Round(fpMatMoraBase.MoraSal / fpMatMoraBase.SaldoSal * 100, 2)
'                End If
'        Case 4:
'                If fpMatMoraBase.Saldo = 0 Then
'                    Resultado = 0
'                Else
'                    Resultado = Round(fpMatMoraBase.Mora / fpMatMoraBase.Saldo * 100, 2)
'                End If
'    End Select
'
'    feMoraDetalle.TextMatrix(i, 5) = Format(Resultado, "###," & String(15, "#") & "#0.00")
'Next i
'
'End Sub
'
'
