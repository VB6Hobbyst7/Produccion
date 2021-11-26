VERSION 5.00
Begin VB.Form frmCapFoncodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Foncodes"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmCapFoncodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6435
      TabIndex        =   17
      Top             =   6990
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3615
      TabIndex        =   16
      Top             =   6990
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4905
      TabIndex        =   15
      Top             =   6990
      Width           =   1170
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   45
      TabIndex        =   10
      Top             =   1290
      Width           =   7590
      Begin SICMACT.FlexEdit feCuotas 
         Height          =   1755
         Left            =   105
         TabIndex        =   46
         Top             =   1695
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   3096
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#---Credito-Cuota-Fec. Venc.-Capital-Interes-Int Comp-Mora-Porte-Protesto-Com Venc-Gastos"
         EncabezadosAnchos=   "0-0-400-0-500-1200-1200-1000-1000-1000-1000-1000-1000-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-L-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-2-2-2-2-2-2-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.EditMoney txtTotal 
         Height          =   315
         Left            =   6120
         TabIndex        =   45
         Top             =   5010
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         BackColor       =   12648447
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboTipoPago 
         Height          =   315
         ItemData        =   "frmCapFoncodes.frx":030A
         Left            =   5850
         List            =   "frmCapFoncodes.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label lblCalendario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1305
         TabIndex        =   44
         Top             =   5535
         Width           =   570
      End
      Begin VB.Label Label12 
         Caption         =   "Detalle"
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
         Height          =   240
         Left            =   150
         TabIndex        =   43
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblCuotasPend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3915
         TabIndex        =   42
         Top             =   810
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "Cuotas Pendien."
         Height          =   225
         Left            =   2610
         TabIndex        =   41
         Top             =   885
         Width           =   1290
      End
      Begin VB.Label Label22 
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4830
         TabIndex        =   40
         Top             =   5070
         Width           =   1185
      End
      Begin VB.Label Label21 
         Caption         =   "Protesto"
         Height          =   270
         Left            =   3105
         TabIndex        =   39
         Top             =   4545
         Width           =   645
      End
      Begin VB.Label lblProtesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3825
         TabIndex        =   38
         Top             =   4485
         Width           =   1245
      End
      Begin VB.Label Label19 
         Caption         =   "Gastos"
         Height          =   270
         Left            =   5460
         TabIndex        =   37
         Top             =   4545
         Width           =   570
      End
      Begin VB.Label lblGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6120
         TabIndex        =   36
         Top             =   4470
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Com. Vencida"
         Height          =   285
         Left            =   195
         TabIndex        =   35
         Top             =   4530
         Width           =   1035
      End
      Begin VB.Label lblComVenc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1485
         TabIndex        =   34
         Top             =   4455
         Width           =   1170
      End
      Begin VB.Label lblCuotasPag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1485
         TabIndex        =   33
         Top             =   3615
         Width           =   600
      End
      Begin VB.Label Label16 
         Caption         =   "Cuotas a Pagar"
         Height          =   255
         Left            =   195
         TabIndex        =   32
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Label Label15 
         Caption         =   "Mora"
         Height          =   270
         Left            =   3090
         TabIndex        =   31
         Top             =   4095
         Width           =   540
      End
      Begin VB.Label lblMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3810
         TabIndex        =   30
         Top             =   4035
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Portes"
         Height          =   270
         Left            =   5445
         TabIndex        =   29
         Top             =   4095
         Width           =   480
      End
      Begin VB.Label lblPortes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6105
         TabIndex        =   28
         Top             =   4020
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Int Compens."
         Height          =   285
         Left            =   210
         TabIndex        =   27
         Top             =   4095
         Width           =   1035
      End
      Begin VB.Label lblIntComp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   26
         Top             =   4035
         Width           =   1170
      End
      Begin VB.Label Label9 
         Caption         =   "Interes"
         Height          =   270
         Left            =   5445
         TabIndex        =   25
         Top             =   3675
         Width           =   600
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6105
         TabIndex        =   24
         Top             =   3585
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Capital"
         Height          =   270
         Left            =   3090
         TabIndex        =   23
         Top             =   3690
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Pago"
         Height          =   225
         Left            =   5010
         TabIndex        =   21
         Top             =   855
         Width           =   825
      End
      Begin VB.Label lblCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3825
         TabIndex        =   20
         Top             =   3615
         Width           =   1245
      End
      Begin VB.Label lblDeuda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   19
         Top             =   4905
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Deuda Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   18
         Top             =   4950
         Width           =   1185
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   885
         TabIndex        =   14
         Top             =   795
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   165
         TabIndex        =   13
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre "
         Height          =   225
         Left            =   165
         TabIndex        =   12
         Top             =   405
         Width           =   705
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   885
         TabIndex        =   11
         Top             =   300
         Width           =   6570
      End
   End
   Begin VB.Frame fraCredito 
      Height          =   1260
      Left            =   60
      TabIndex        =   4
      Top             =   15
      Width           =   7575
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4755
         TabIndex        =   8
         Top             =   165
         Width           =   2685
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCapFoncodes.frx":030E
            Left            =   75
            List            =   "frmCapFoncodes.frx":0310
            TabIndex        =   9
            Top             =   225
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3765
         TabIndex        =   7
         Top             =   555
         Width           =   840
      End
      Begin VB.Frame FraCta 
         Height          =   585
         Left            =   150
         TabIndex        =   5
         Top             =   360
         Width           =   3420
         Begin VB.TextBox TxtCta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   885
            MaxLength       =   3
            TabIndex        =   0
            Top             =   165
            Width           =   450
         End
         Begin VB.TextBox TxtCta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1335
            MaxLength       =   2
            TabIndex        =   1
            Top             =   165
            Width           =   345
         End
         Begin VB.TextBox TxtCta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1710
            MaxLength       =   7
            TabIndex        =   2
            Top             =   165
            Width           =   870
         End
         Begin VB.TextBox TxtCta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2580
            MaxLength       =   6
            TabIndex        =   3
            Top             =   165
            Width           =   765
         End
         Begin VB.Label label1 
            Caption         =   "Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   690
         End
      End
   End
End
Attribute VB_Name = "frmCapFoncodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbVencido As Boolean

Private PlanPagos As Variant

Dim lsCodCta As String * 18

Private Sub cboTipoPago_Click()
Dim i As Integer

    If cboTipoPago.Text <> "" Then
        feCuotas.Enabled = True
        If cboTipoPago.ListIndex = 1 Then
            For i = 0 To feCuotas.Rows - 1
                feCuotas.TextMatrix(i, 2) = 1
                feCuotas_OnCellCheck 1, 2
            Next i
        Else
            For i = 0 To feCuotas.Rows - 1
                feCuotas.TextMatrix(i, 2) = 0
                feCuotas_OnCellCheck 1, 2
            Next i
        End If
    End If

End Sub

Private Sub cboTipoPago_KeyPress(KeyAscii As Integer)
Dim i As Integer
    If KeyAscii = 13 Then
        If cboTipoPago.Text <> "" Then
            feCuotas.Enabled = True
            If cboTipoPago.ListIndex = 1 Then
                For i = 0 To feCuotas.Rows - 1
                    feCuotas.TextMatrix(i, 2) = 1
                    feCuotas_OnCellCheck 1, 2
                Next i
            End If
        End If
    End If

End Sub

Private Sub cmdBuscar_Click()

    frmBuscaPersFioncodes.Inicio

End Sub

Private Sub cmdCancelar_Click()
    Blanquea
End Sub


Private Sub cmdGrabar_Click()
Dim oGraba As NCapFideicomiso
Dim oContFunc As NContFunciones
Dim lsMovNro As String
Dim lnCapital As Currency
Dim lnInteres As Currency
Dim lnIntComp As Currency
Dim lnMora As Currency
Dim lnPortes As Currency
Dim lnProtesto As Currency
Dim lnComVcdo As Currency
Dim lnGastos As Currency
Dim lnTCapital As Currency
Dim lnTInteres As Currency
Dim lnTIntComp As Currency
Dim lnTMora As Currency
Dim lnTPortes As Currency
Dim lnTProtesto As Currency
Dim lnTComVcdo As Currency
Dim lnTGastos As Currency
Dim lnCuotasPag As Integer
Dim lnPago As Currency
Dim i As Integer
Dim lnCuotaPagPJ As Integer

lnCapital = 0:      lnInteres = 0:      lnIntComp = 0:      lnMora = 0:     lnPortes = 0
lnProtesto = 0:     lnComVcdo = 0:      lnGastos = 0

lnTCapital = 0:      lnTInteres = 0:      lnTIntComp = 0:      lnTMora = 0:     lnTPortes = 0
lnTProtesto = 0:     lnTComVcdo = 0:      lnTGastos = 0

'Falta verificar si se puede realizar el pago

If MsgBox("Se va a Efectuar el Pago del FONCODES, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then

    If lblCalendario = 0 Then   'NORMAL
        ReDim PlanPagos(1000, 9)
        For i = 1 To feCuotas.Rows - 1
            If feCuotas.TextMatrix(i, 2) = "." Then
                If cboTipoPago.ListIndex = 0 Then
                    lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                    lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                    lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                    lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                    lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                    lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                    lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                    lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                Else
                    If lbVencido Then
                        If DateDiff("d", feCuotas.TextMatrix(i, 5), gdFecSis) > 0 Then
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                            lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                            lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                            lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                            lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                            lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                            lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                            lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                        Else
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                        End If
                    Else
                        If i = 1 Then
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                            lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                            lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                            lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                            lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                            lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                            lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                            lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                        Else
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                        End If
                    End If
                End If

                PlanPagos(i, 1) = feCuotas.TextMatrix(i, 4)
                PlanPagos(i, 2) = feCuotas.TextMatrix(i, 6)
                PlanPagos(i, 3) = feCuotas.TextMatrix(i, 7)
                PlanPagos(i, 4) = feCuotas.TextMatrix(i, 8)
                PlanPagos(i, 5) = feCuotas.TextMatrix(i, 9)
                PlanPagos(i, 6) = feCuotas.TextMatrix(i, 10)
                PlanPagos(i, 7) = feCuotas.TextMatrix(i, 11)
                PlanPagos(i, 8) = feCuotas.TextMatrix(i, 12)
                PlanPagos(i, 9) = feCuotas.TextMatrix(i, 13)

            End If
        Next i

        'Totales para la impresion
        lnTCapital = lnCapital
        lnTInteres = lnInteres
        lnTPortes = lnPortes
        lnTIntComp = lnIntComp
        lnTMora = lnMora
        lnTProtesto = lnProtesto
        lnTComVcdo = lnComVcdo
        lnTGastos = lnGastos

        lnCuotasPag = lblCuotasPag

    Else    'PREJUDICIAL

       i = 1
       If CCur(lblDeuda) > CCur(txtTotal) Then
          lnCuotaPagPJ = 0
       Else
          lnCuotaPagPJ = 1
       End If
       lnPago = txtTotal

       ReDim PlanPagos(1, 9)
        PlanPagos(1, 1) = feCuotas.TextMatrix(i, 4)

        If lnPago > 0 Then  'Gastos
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 13)) Then
                PlanPagos(1, 9) = feCuotas.TextMatrix(i, 13)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 13))
                lnTGastos = feCuotas.TextMatrix(i, 13)
            Else
                PlanPagos(1, 9) = lnPago
                lnTGastos = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then  'Protesto
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 11)) Then
                PlanPagos(1, 7) = feCuotas.TextMatrix(i, 11)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 11))
                lnTProtesto = feCuotas.TextMatrix(i, 11)
            Else
                PlanPagos(1, 7) = lnPago
                lnTProtesto = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then  'Com Vcdo
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 12)) Then
                PlanPagos(1, 8) = feCuotas.TextMatrix(i, 12)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 12))
                lnTComVcdo = feCuotas.TextMatrix(i, 12)
            Else
                PlanPagos(1, 8) = lnPago
                lnTComVcdo = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then      'Mora
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 9)) Then
                PlanPagos(1, 5) = feCuotas.TextMatrix(i, 9)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 9))
                lnTMora = feCuotas.TextMatrix(i, 9)
            Else
                PlanPagos(1, 5) = lnPago
                lnTMora = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then      'Int Comp
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 8)) Then
                PlanPagos(1, 4) = feCuotas.TextMatrix(i, 8)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 8))
                lnTIntComp = feCuotas.TextMatrix(i, 8)
            Else
                PlanPagos(1, 4) = lnPago
                lnTIntComp = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then      'Portes
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 10)) Then
                PlanPagos(1, 6) = feCuotas.TextMatrix(i, 10)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 10))
                lnTPortes = feCuotas.TextMatrix(i, 10)
            Else
                PlanPagos(1, 6) = lnPago
                lnTPortes = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then  'Interes
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 7)) Then
                PlanPagos(1, 3) = feCuotas.TextMatrix(i, 7)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 7))
                lnTInteres = feCuotas.TextMatrix(i, 7)
            Else
                PlanPagos(1, 3) = lnPago
                lnTInteres = lnPago
                lnPago = 0
            End If
        End If

        If lnPago > 0 Then  'Capital
            If CCur(lnPago) > CCur(feCuotas.TextMatrix(i, 6)) Then
                PlanPagos(1, 2) = feCuotas.TextMatrix(i, 6)
                lnPago = CCur(lnPago) - CCur(feCuotas.TextMatrix(i, 6))
                lnTCapital = feCuotas.TextMatrix(i, 6)
            Else
                PlanPagos(1, 2) = lnPago
                lnTCapital = lnPago
                lnPago = 0
            End If
        End If

        lnCuotasPag = 1

    End If

    Set oContFunc = New NContFunciones
        lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oContFunc = Nothing

    Set oGraba = New NCapFideicomiso
    oGraba.CapPagoFONCODES lsCodCta, PlanPagos, txtTotal, lnCuotasPag, cboTipoPago.ListIndex, lsMovNro, lblCalendario, lnCuotaPagPJ

    oGraba.ImprimeBoleta lsCodCta, lblNombre, gsNomAge, lblMoneda, lblCuotasPag, gdFecSis, Time, lnTCapital, lnTInteres, _
                         lnTPortes, lnTIntComp, lnTMora, lnTComVcdo, lnTProtesto, lnTGastos, gsCodUser, sLpt

    Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes

        oGraba.ImprimeBoleta lsCodCta, lblNombre, gsNomAge, lblMoneda, lblCuotasPag, gdFecSis, Time, lnTCapital, lnTInteres, _
                             lnTPortes, lnTIntComp, lnTMora, lnTComVcdo, lnTProtesto, lnTGastos, gsCodUser, sLpt

    Loop

    Set oGraba = Nothing
    Blanquea

End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub feCuotas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim lnCapital As Currency
Dim lnInteres As Currency
Dim lnIntComp As Currency
Dim lnMora As Currency
Dim lnPortes As Currency
Dim lnProtesto As Currency
Dim lnComVcdo As Currency
Dim lnGastos As Currency
Dim lnCuotasPag As Integer
Dim i As Integer

Screen.MousePointer = 11

lnCapital = 0:      lnInteres = 0:      lnIntComp = 0:      lnMora = 0:     lnPortes = 0
lnProtesto = 0:     lnComVcdo = 0:      lnGastos = 0:       lnCuotasPag = 0

    For i = 1 To feCuotas.Rows - 1

        If feCuotas.TextMatrix(i, 2) = "." Then

            If cboTipoPago.ListIndex = 0 Then   'AMORTIZACION
                lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
            Else        'CANCELACION

                If lblCalendario = 1 Then   'NORMAL

                    If lbVencido Then       'CUOTAS VENCIDAS
                        If DateDiff("d", feCuotas.TextMatrix(i, 5), gdFecSis) > 0 Then
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                            lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                            lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                            lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                            lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                            lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                            lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                            lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                        Else                'CUOTA NO VENCIDA
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                        End If

                    Else
                        If i = 1 Then       'PRIMERA CUOTA
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                            lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                            lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                            lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                            lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                            lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                            lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                            lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                        Else                'DEMAS CUOTAS
                            lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                        End If
                    End If

                Else        'PREJUDICIAL
                    lnCapital = lnCapital + CCur(feCuotas.TextMatrix(i, 6))
                    lnInteres = lnInteres + CCur(feCuotas.TextMatrix(i, 7))
                    lnIntComp = lnIntComp + CCur(feCuotas.TextMatrix(i, 8))
                    lnMora = lnMora + CCur(feCuotas.TextMatrix(i, 9))
                    lnPortes = lnPortes + CCur(feCuotas.TextMatrix(i, 10))
                    lnProtesto = lnProtesto + CCur(feCuotas.TextMatrix(i, 11))
                    lnComVcdo = lnComVcdo + CCur(feCuotas.TextMatrix(i, 12))
                    lnGastos = lnGastos + CCur(feCuotas.TextMatrix(i, 13))
                End If

            End If

            lnCuotasPag = lnCuotasPag + 1

        End If
    Next i

    lblCapital = Format(lnCapital, "######,###.00")
    lblInteres = Format(lnInteres, "###,###.00")
    lblIntComp = Format(lnIntComp, "###,###.00")
    lblMora = Format(lnMora, "###,###.00")
    lblPortes = Format(lnPortes, "###,###.00")
    lblProtesto = Format(lnProtesto, "###,###.00")
    lblComVenc = Format(lnComVcdo, "###,###.00")
    lblGastos = Format(lnGastos, "###,###.00")
    txtTotal = lnCapital + lnInteres + lnIntComp + lnMora + lnPortes + lnProtesto + lnComVcdo + lnGastos
    lblCuotasPag = lnCuotasPag

Screen.MousePointer = 1

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
    cboTipoPago.AddItem "AMORTIZACION", 0
    cboTipoPago.AddItem "CANCELACION", 1

End Sub

Private Sub LstCred_Click()

    If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then

        TxtCta(0).Text = Mid(LstCred.Text, 1, 3)
        TxtCta(1).Text = Mid(LstCred.Text, 4, 2)
        TxtCta(2).Text = Mid(LstCred.Text, 6, 7)
        TxtCta(3).Text = Mid(LstCred.Text, 13, 6)

    End If

End Sub

Private Sub TxtCta_Change(Index As Integer)

    Select Case Index
    Case 0
        If Len(TxtCta(0)) = 3 Then
            TxtCta(1).SetFocus
        Else
            TxtCta(0).SetFocus
        End If
    Case 1
        If Len(TxtCta(1)) = 2 Then
            TxtCta(2).SetFocus
        Else
            TxtCta(1).SetFocus
        End If
    Case 2
        If Len(TxtCta(2)) = 7 Then
            TxtCta(3).SetFocus
        Else
            TxtCta(2).SetFocus
        End If
    End Select

End Sub

Private Sub TxtCta_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        Select Case Index
        Case 0
            If Len(TxtCta(0).Text) = 3 Then
                TxtCta(1).SetFocus
            Else
                TxtCta(0).SetFocus
            End If
        Case 1
            If Len(TxtCta(1).Text) = 2 Then
                TxtCta(2).SetFocus
            Else
                TxtCta(1).SetFocus
            End If
        Case 2
            If Len(TxtCta(2).Text) = 7 Then
                TxtCta(3).SetFocus
            Else
                TxtCta(2).SetFocus
            End If
        Case 3
            If Len(TxtCta(3).Text) = 6 Then
                lsCodCta = Trim(TxtCta(0)) & Trim(TxtCta(1)) & Trim(TxtCta(2)) & Trim(TxtCta(3))
                If Len(lsCodCta) = 18 Then
                    BuscaFoncodes lsCodCta
                    cmdGrabar.Enabled = True
                Else
                    MsgBox "Nro de Cuenta no válida, Verifique el dato", vbInformation, "Aviso"
                End If
            Else
                TxtCta(3).SetFocus
            End If
        End Select
    ElseIf KeyAscii = vbKeyBack Then
        Select Case Index
        Case 1
            If Len(TxtCta(1)) = 0 Then
                TxtCta(0).SetFocus
            End If
        Case 2
            If Len(TxtCta(2)) = 0 Then
                TxtCta(1).SetFocus
            End If
        Case 3
            If Len(TxtCta(3)) = 0 Then
                TxtCta(2).SetFocus
            End If
        End Select

    End If

End Sub

Private Sub BuscaFoncodes(ByVal psCodCta As String)
Dim oDatos As DCapFideicomiso
Dim rs As Recordset
Dim lnInteres As Currency
Dim lnGastos As Currency

Set oDatos = New DCapFideicomiso

Set rs = oDatos.dGetDatosFoncodes(psCodCta)

    If Not rs.EOF And Not rs.BOF Then
        If rs!nNroCuotasPend > 0 Then
            lblNombre = rs!cNombre
            lblDeuda = rs!nMontoTotalDeuda
            lblCalendario = rs!nIndCalendario
            lblCuotasPend = rs!nNroCuotasPend
            lblMoneda = IIf(Right(TxtCta(1), 1) = "1", "SOLES", "DOLARES")
            fraCredito.Enabled = False
            FraDatos.Enabled = True
            If rs!nIndCalendario = 0 Then
                txtTotal.Enabled = False
            Else
                txtTotal.Enabled = True
            End If
        Else
            MsgBox "Cuenta no posee cuotas pendientes", vbInformation, "Aviso"
            Blanquea
            Exit Sub
        End If
    Else
        MsgBox "Cuenta no válida", vbInformation, "Aviso"
        Blanquea
        Exit Sub
    End If

Set rs = Nothing

Set rs = oDatos.dGetDatosFoncodesDet(psCodCta)

    If Not rs.EOF And Not rs.BOF Then

        lbVencido = False
        If DateDiff("d", Format(rs!dFecVenc, "dd/mm/yyyy"), gdFecSis) > 0 Then
            lbVencido = True
        End If

        Do While Not rs.EOF
            feCuotas.AdicionaFila
            feCuotas.TextMatrix(feCuotas.Rows - 1, 1) = feCuotas.Rows - 1
            feCuotas.TextMatrix(feCuotas.Rows - 1, 3) = psCodCta
            feCuotas.TextMatrix(feCuotas.Rows - 1, 4) = rs!nNroCuota
            feCuotas.TextMatrix(feCuotas.Rows - 1, 5) = Format(rs!dFecVenc, "dd/mm/yyyy")
            feCuotas.TextMatrix(feCuotas.Rows - 1, 6) = rs!nCapital
            feCuotas.TextMatrix(feCuotas.Rows - 1, 7) = rs!nInteres
            feCuotas.TextMatrix(feCuotas.Rows - 1, 8) = rs!nInteresComp
            feCuotas.TextMatrix(feCuotas.Rows - 1, 9) = rs!nMora
            feCuotas.TextMatrix(feCuotas.Rows - 1, 10) = rs!nPortes
            feCuotas.TextMatrix(feCuotas.Rows - 1, 11) = rs!nProtesto
            feCuotas.TextMatrix(feCuotas.Rows - 1, 12) = rs!nComVcdo
            feCuotas.TextMatrix(feCuotas.Rows - 1, 13) = rs!nGastos
            rs.MoveNext
        Loop
    End If

Set rs = Nothing
Set oDatos = Nothing

End Sub

Private Sub Blanquea()

    lblComVenc = ""
    lblCuotasPag = ""
    lblCapital = ""
    lblDeuda = ""
    lblIntComp = ""
    lblGastos = ""
    lblInteres = ""
    lblMoneda = ""
    lblNombre = ""
    lblMora = ""
    lblPortes = ""
    lblProtesto = ""
    txtTotal.Text = ""
    feCuotas.Clear
    feCuotas.Rows = 2
    feCuotas.FormaCabecera
    fraCredito.Enabled = True
    FraDatos.Enabled = False
    feCuotas.Enabled = False
    FraCta.Enabled = True
    TxtCta(0).Text = ""
    TxtCta(1).Text = ""
    TxtCta(2).Text = ""
    TxtCta(3).Text = ""
    LstCred.Clear
    cboTipoPago.ListIndex = -1
    cmdGrabar.Enabled = False

End Sub
