VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogAlmCodBarra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Código de Barras"
   ClientHeight    =   5235
   ClientLeft      =   1605
   ClientTop       =   1905
   ClientWidth     =   8430
   Icon            =   "frmLogAlmCodBarra.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8430
   Begin VB.CommandButton cmdGenera 
      Caption         =   "Generar"
      Height          =   375
      Left            =   7080
      TabIndex        =   30
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   2700
      TabIndex        =   17
      Top             =   2940
      Width           =   4275
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "0000"
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   780
         MaxLength       =   8
         TabIndex        =   31
         Text            =   "00000000"
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   240
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "2006"
         Top             =   300
         Width           =   495
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         Picture         =   "frmLogAlmCodBarra.frx":08CA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   24
         Top             =   900
         Width           =   375
      End
      Begin VB.PictureBox picBarCode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   23
         Top             =   1260
         Width           =   3015
      End
      Begin VB.CommandButton cmdCodBarra 
         Caption         =   "Código de Barras"
         Height          =   615
         Left            =   3060
         TabIndex        =   22
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIO GENERAL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   28
         Top             =   1740
         Width           =   2145
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRUJILLO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   27
         Top             =   1560
         Width           =   2070
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRUJILLO"
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   26
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   25
         Top             =   960
         Width           =   345
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1275
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   780
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2115
      Left            =   120
      TabIndex        =   14
      Top             =   2940
      Width           =   2415
      Begin VB.TextBox txtRX 
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtRY 
         Height          =   315
         Left            =   1260
         TabIndex        =   18
         Top             =   1080
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1260
         TabIndex        =   15
         Text            =   "20"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desplaz. Y"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desplaz. X"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Alto Barras"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8175
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   7695
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   7695
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1110
         Width           =   7695
      End
      Begin VB.TextBox txtDescripcion 
         BackColor       =   &H00EAFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   5
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1380
         Width           =   7695
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   570
         Width           =   7695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Descripción del Bien "
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
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8175
      Begin VB.TextBox txtBSCod 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Index           =   0
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Left            =   4020
         TabIndex        =   13
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Bien"
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
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   1350
      End
   End
   Begin MSComDlg.CommonDialog comPrinter 
      Left            =   180
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmLogAlmCodBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim cBSCod As String, cBSDescripcion As String, cBSSerie As String
Dim nTamanio As Integer, nAlto As Integer

Public Sub Codigo(ByVal psBSCod As String, ByVal psBSDescripcion As String, Optional psBSSerie As String = "")
cBSCod = psBSCod
cBSDescripcion = psBSDescripcion
cBSSerie = psBSSerie
Me.Show 1
End Sub


Private Sub cmdGenera_Click()
Dim lsCodigoBarra As String
Dim k As Integer, i As Integer
Dim P As Printer
Dim OK As Boolean

OK = False
For Each P In Printers
   If InStr(UCase(P.DeviceName), "ZEBRA") > 0 Then
      OK = True
      Exit For
   End If
Next

k = 1050
If OK Then
   'For i = 1 To 50 / 2
 For i = 1 To 1
       k = k + 1
       lsCodigoBarra = ""
       txtCodigo.Text = Format(k, "00000000"):  DibujaBarCode
       
       Printer.Font.Name = "Helvetica": Printer.Font.Size = 5: Printer.Font.Bold = True
       Printer.CurrentX = 800: Printer.CurrentY = 310:   Printer.Print "CAJA"
       Printer.CurrentX = 800: Printer.CurrentY = 410:   Printer.Print "TRUJILLO"
       Printer.PaintPicture picLogo, 440, 230
       
       Printer.Font.Name = "Arial": Printer.Font.Size = 8.16: Printer.Font.Bold = False
       SavePicture picBarCode.Image, App.Path & "\Temp.bmp":      Sleep 300
       picBarCode.Picture = LoadPicture(App.Path & "\Temp.bmp"):  Sleep 300
       Printer.PaintPicture picBarCode, 140, 570
       
       Printer.Font.Name = "Arial Narrow":  Printer.Font.Size = 7:  Printer.Font.Bold = True
       Printer.CurrentX = 500:  Printer.CurrentY = 820:   Printer.Print "INVENTARIO CMACT " + txtAnio.Text & " " & txtCodigo
       Set picBarCode.Picture = LoadPicture("")
       
       k = k + 1
       txtCodigo.Text = Format(k, "00000000")
       
       DibujaBarCode
       Printer.Font.Bold = True
       Printer.Font.Name = "Helvetica"
       Printer.Font.Size = 5
       Printer.CurrentX = 3900:  Printer.CurrentY = 310:   Printer.Print "CAJA"
       Printer.CurrentX = 3900:  Printer.CurrentY = 410:   Printer.Print "TRUJILLO"
       Printer.PaintPicture picLogo, 3540, 230
       Printer.DrawWidth = 1
       
       Printer.Font.Name = "Arial":  Printer.Font.Size = 8.16:   Printer.Font.Bold = False
       SavePicture picBarCode.Image, App.Path & "\Temp.bmp":     Sleep 300
       picBarCode.Picture = LoadPicture(App.Path & "\Temp.bmp"): Sleep 300
       Printer.PaintPicture picBarCode, 3240, 570
       
       Printer.Font.Name = "Arial Narrow":  Printer.Font.Size = 7:  Printer.Font.Bold = True
       Printer.CurrentX = 3600:  Printer.CurrentY = 820:   Printer.Print "INVENTARIO CMACT " + txtAnio.Text & " " & txtCodigo
       Set picBarCode.Picture = LoadPicture("")
        
       Printer.EndDoc
   Next
Else
   MsgBox "No hay una impresora de Etiquetas instalada..." + Space(10), vbInformation, "Aviso"
End If



End Sub

Private Sub cmdPrint_Click()
Dim P As Printer
Dim OK As Boolean
Dim k As Integer

OK = False
For Each P In Printers
   If InStr(UCase(P.DeviceName), "ZEBRA") > 0 Then
      OK = True
      Exit For
   End If
Next

If OK Then
   Printer.Font.Bold = True
   Printer.Font.Name = "Helvetica"
   Printer.Font.Size = 5
   Printer.CurrentY = 310: Printer.CurrentX = 880:  Printer.Print "CAJA"
   Printer.CurrentY = 410: Printer.CurrentX = 880:  Printer.Print "TRUJILLO"
   Printer.PaintPicture picLogo, 520, 230
   Printer.DrawWidth = 1
   
   Printer.Font.Bold = False
   Printer.Font.Name = "Arial"
   Printer.Font.Size = 8.16
   
   SavePicture picBarCode.Image, App.Path & "\Temp.bmp"
   Sleep 300
   picBarCode.Picture = LoadPicture(App.Path & "\Temp.bmp")
   Sleep 300

   Printer.PaintPicture picBarCode, 520, 570
   
   Printer.Font.Bold = True
   Printer.Font.Name = "Arial Narrow"
   Printer.Font.Size = 7
   Printer.CurrentY = 810: Printer.CurrentX = 680:   Printer.Print "2006-1234567890123456-0000"
   'Printer.Font.Size = 5
   'Printer.CurrentY = 960: Printer.CurrentX = 940:   Printer.Print "INVENTARIO GENERAL"
   
   Printer.Font.Bold = True
   Printer.Font.Name = "Helvetica"
   Printer.Font.Size = 5
   Printer.CurrentY = 310: Printer.CurrentX = 3980:  Printer.Print "CAJA"
   Printer.CurrentY = 410: Printer.CurrentX = 3980:  Printer.Print "TRUJILLO"
   Printer.PaintPicture picLogo, 3620, 230
   Printer.DrawWidth = 1
   
   Printer.Font.Bold = False
   Printer.Font.Name = "Arial"
   Printer.Font.Size = 8.16
   Printer.PaintPicture picBarCode, 3620, 570
   
   Printer.Font.Bold = True
   Printer.Font.Name = "Arial Narrow"
   Printer.Font.Size = 7
   Printer.CurrentY = 810: Printer.CurrentX = 3780:   Printer.Print "00002005-0123456789-0123"
   Printer.Font.Size = 5
   Printer.CurrentY = 960: Printer.CurrentX = 4080:   Printer.Print "INVENTARIO GENERAL"
   Printer.EndDoc
Else
   MsgBox "No hay una impresora de Etiquetas instalada..." + Space(10), vbInformation, "Aviso"
End If
End Sub


Private Sub Form_Load()
txtBSCod = cBSCod
txtDescripcion(5).Text = cBSDescripcion
txtDescripcion(0).Text = cBSSerie
'DescripcionBien txtBSCod
txtRX.Text = 100
txtRY.Text = 100
'DibujaBarCode cBSCod, picBarCode, 1, 15
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub Text3_Change()

End Sub

Private Sub txtBSCod_Change()
Dim i As Integer, cBSCod As String
If Len(Trim(txtBSCod.Text)) >= 8 Then
   'DibujaBarCode txtBSCod.Text, picBarCode, 1, 15
Else
   Set picBarCode.Picture = LoadPicture("")
End If
End Sub

'Sub DescripcionBien(ByVal psBSCod As String)
'Dim rs As New ADODB.Recordset
'Dim oConn As New DConecta
'Dim i As Integer
'Dim sSQL As String
'
'i = 0
'sSQL = "select cBSCod,cBSDescripcion from BienesServicios where cBSCod = '" & Left(psBSCod, 2) & "' union " & _
'       "select cBSCod,cBSDescripcion from BienesServicios where cBSCod = '" & Left(psBSCod, 3) & "' union " & _
'       "select cBSCod,cBSDescripcion from BienesServicios where cBSCod = '" & Left(psBSCod, 5) & "' union " & _
'       "select cBSCod,cBSDescripcion from BienesServicios where cBSCod = '" & Left(psBSCod, 8) & "' union " & _
'       "select cBSCod,cBSDescripcion from BienesServicios where cBSCod = '" & psBSCod & "' " & _
'       "order by cBSCod "
'If oConn.AbreConexion Then
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      Do While Not rs.EOF
'         i = i + 1
'         txtDescripcion(i).Text = rs!cBSDescripcion
'         rs.MoveNext
'      Loop
'   End If
'End If
'End Sub

Private Sub cmdCodBarra_Click()
'DibujaBarCode txtCodigo, picBarCod, 1
DibujaBarCode
End Sub


Private Sub DibujaBarCode()
Dim X As Integer, Y As Integer, z As Integer, pos As Integer
Dim Bardata As String
Dim Cur As String
Dim CurVal As Long
Dim chksum As Long
Dim chkchr As String
Dim BC(43) As String
    '3 of the 9 bars are wide: 0=narrow, 1=wide
    BC(0) = "000110100" '0
    BC(1) = "100100001" '1
    BC(2) = "001100001" '2
    BC(3) = "101100000" '3
    BC(4) = "000110001" '4
    BC(5) = "100110000" '5
    BC(6) = "001110000" '6
    BC(7) = "000100101" '7
    BC(8) = "100100100" '8
    BC(9) = "001100100" '9
    BC(10) = "100001001" 'A
    BC(11) = "001001001" 'B
    BC(12) = "101001000" 'C
    BC(13) = "000011001" 'D
    BC(14) = "100011000" 'E
    BC(15) = "001011000" 'F
    BC(16) = "000001101" 'G
    BC(17) = "100001100" 'H
    BC(18) = "001001100" 'I
    BC(19) = "000011100" 'J
    BC(20) = "100000011" 'K
    BC(21) = "001000011" 'L
    BC(22) = "101000010" 'M
    BC(23) = "000010011" 'N
    BC(24) = "100010010" 'O
    BC(25) = "001010010" 'P
    BC(26) = "000000111" 'Q
    BC(27) = "100000110" 'R
    BC(28) = "001000110" 'S
    BC(29) = "000010110" 'T
    BC(30) = "110000001" 'U
    BC(31) = "011000001" 'V
    BC(32) = "111000000" 'W
    BC(33) = "010010001" 'X
    BC(34) = "110010000" 'Y
    BC(35) = "011010000" 'Z
    BC(36) = "010000101" '-
    BC(37) = "110000100" '.
    BC(38) = "011000100" '<spc>
    BC(39) = "010101000" '$
    BC(40) = "010100010" '/
    BC(41) = "010001010" '+
    BC(42) = "000101010" '%
    BC(43) = "010010100" '*  (used for start/stop character only)
    
picBarCode.Cls
If txtCodigo.Text = "" Then Exit Sub
pos = 20
Bardata = UCase(txtCodigo.Text)

'Check for invalid characters and calculate check sum
For X = 1 To Len(Bardata)
    Cur = Mid$(Bardata, X, 1)
    Select Case Cur
        Case "0" To "9"
            CurVal = Val(Cur)
        Case "A" To "Z"
            CurVal = Asc(Cur) - 55
        Case "-"
            CurVal = 36
        Case "."
            CurVal = 37
        Case " "
            CurVal = 38
        Case "$"
            CurVal = 39
        Case "/"
            CurVal = 40
        Case "+"
            CurVal = 41
        Case "%"
            CurVal = 42
        Case Else 'oops!
            picBarCode.Cls
            picBarCode.Print Cur & " is Invalid"
            Exit Sub
    End Select
    chksum = chksum + CurVal
Next

'Aqui la etiqueta
'If Check1(1).Value Then
'    picBarCode.CurrentX = 35 + Len(Bardata) * (5)   'kinda center
'    picBarCode.CurrentY = 50
'    picBarCode.Print Bardata;
'End If

'Add Check Character? (rarely used, but i put it here anyway...)
'If Check1(2).value Then
'    chksum = chksum Mod 43
'    chkchr = Mid$("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*", chksum + 1, 1)
'    picBarCode.Print "_" & chkchr
'    Bardata = Bardata & chkchr
'End If

'Add Start & Stop characters (must have 'em for valid barcodes)
Bardata = "*" & Bardata & "*"

'Generate Barcode
For X = 1 To Len(Bardata)
    Cur = Mid$(Bardata, X, 1)
    Select Case Cur
        Case "0" To "9"
            CurVal = Val(Cur)
        Case "A" To "Z"
            CurVal = Asc(Cur) - 55
        Case "-"
            CurVal = 36
        Case "."
            CurVal = 37
        Case " "
            CurVal = 38
        Case "$"
            CurVal = 39
        Case "/"
            CurVal = 40
        Case "+"
            CurVal = 41
        Case "%"
            CurVal = 42
        Case "*"
            CurVal = 43
    End Select
    
    For Y = 1 To 9
        If Y Mod 2 = 0 Then
            'SPACE
            pos = pos + 1 + (2 * Val(Mid$(BC(CurVal), Y, 1)))
        Else
            'BAR
            For z = 1 To 1 + (Val(Mid$(BC(CurVal), Y, 1)))
                picBarCode.Line (pos, 1)-(pos, 58 - 40)
                pos = pos + 1
            Next z
        End If
    Next
    pos = pos + 1
Next
End Sub






'Sub DibujaBarCode(ByVal bc_string As String, obj As Control, pnTamanio As Integer, Optional pnAlto As Integer = 15)
'Dim i As Integer, n As Integer
'Dim xpos!, Y1!, Y2!, dw%, th!, tw, new_string$
'Dim c As Integer, bc_pattern As String
'Dim bc(90) As String
'
'nAlto = pnAlto
'nTamanio = pnTamanio
'bc(1) = "1 1221"            'pre-amble
'bc(2) = "1 1221"            'post-amble
'bc(48) = "11 221"           'digitos
'bc(49) = "21 112"
'bc(50) = "12 112"
'bc(51) = "22 111"
'bc(52) = "11 212"
'bc(53) = "21 211"
'bc(54) = "12 211"
'bc(55) = "11 122"
'bc(56) = "21 121"
'bc(57) = "12 121"
'                            'Mayusculas
'bc(65) = "211 12"           'A
'bc(66) = "121 12"           'B
'bc(67) = "221 11"           'C
'bc(68) = "112 12"           'D
'bc(69) = "212 11"           'E
'bc(70) = "122 11"           'F
'bc(71) = "111 22"           'G
'bc(72) = "211 21"           'H
'bc(73) = "121 21"           'I
'bc(74) = "112 21"           'J
'bc(75) = "2111 2"           'K
'bc(76) = "1211 2"           'L
'bc(77) = "2211 1"           'M
'bc(78) = "1121 2"           'N
'bc(79) = "2121 1"           'O
'bc(80) = "1221 1"           'P
'bc(81) = "1112 2"           'Q
'bc(82) = "2112 1"           'R
'bc(83) = "1212 1"           'S
'bc(84) = "1122 1"           'T
'bc(85) = "2 1112"           'U
'bc(86) = "1 2112"           'V
'bc(87) = "2 2111"           'W
'bc(88) = "1 1212"           'X
'bc(89) = "2 1211"           'Y
'bc(90) = "1 2211"           'Z
'                            'Misc
'bc(32) = "1 2121"           'space
'bc(35) = ""                 '# cannot do!
'bc(36) = "1 1 1 11"         '$
'bc(37) = "11 1 1 1"         '%
'bc(43) = "1 11 1 1"         '+
'bc(45) = "1 1122"           '-
'bc(47) = "1 1 11 1"         '/
'bc(46) = "2 1121"           '.
'bc(64) = ""                 '@ cannot do!
'bc(65) = "1 1221"           '*
'
'bc_string = UCase(bc_string)
'
''dimensiones
'obj.ScaleMode = 3                               'pixels
'obj.Cls
'obj.Picture = Nothing
'dw = nTamanio                 'espacio entre barras
'
''If dw < 1 Then dw = 1         'si no se ha definido espacio entre barras
'dw = 1
'
'th = obj.Height
'Y1 = 2
'Y2 = nAlto
'new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
''dibuja cada caracter de la cadena de entrada
'xpos = obj.ScaleLeft
'For n = 1 To Len(new_string)
'    c = Asc(Mid$(new_string, n, 1))
'    If c > 90 Then c = 0
'    bc_pattern$ = bc(c)
'    'draw each bar
'    For i = 1 To Len(bc_pattern$)
'        Select Case Mid$(bc_pattern$, i, 1)
'            Case " "
'                'space
'                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, B
'                xpos = xpos + dw
'
'            Case "1"
'                'space
'                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, B
'                xpos = xpos + dw
'                'line
'                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, B
'                xpos = xpos + dw
'
'            Case "2"
'                'space
'                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, B
'                xpos = xpos + dw
'                'wide line
'                obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, B
'                xpos = xpos + 2 * dw
'        End Select
'    Next
'Next
''1 more space
'obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
'xpos = xpos + dw
''ancho final1
'obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
''copy to clipboard
'obj.Picture = obj.Image
'Clipboard.Clear
'Clipboard.SetData obj.Image, 2
'End Sub
'
