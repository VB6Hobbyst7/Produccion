VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Object = "{94F82BCB-2A21-4F75-916D-DE398F3D6CA5}#2.0#0"; "PryImage.ocx"
Begin VB.Form frmPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personas: Mantenimiento"
   ClientHeight    =   7050
   ClientLeft      =   1290
   ClientTop       =   2445
   ClientWidth     =   11385
   Icon            =   "PersonasMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdPersCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   9945
      TabIndex        =   66
      ToolTipText     =   "Cancelar Todos los cambios Realizados"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdPersAceptar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
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
      Left            =   8445
      TabIndex        =   65
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin Sicmact.TxtBuscar TxtBCodPers 
      Height          =   330
      Left            =   690
      TabIndex        =   26
      Top             =   90
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   3
      sTitulo         =   ""
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   300
      Left            =   10320
      TabIndex        =   36
      Top             =   210
      Width           =   1000
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   300
      Left            =   10320
      TabIndex        =   37
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   300
      Left            =   10320
      TabIndex        =   38
      Top             =   975
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTDatosGen 
      Height          =   3855
      Left            =   90
      TabIndex        =   47
      Top             =   3120
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "PersonasMantenimiento.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNac"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersCIIU"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbPersEstado"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CboPersCiiu"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPersNacCreac"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CmbRela"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtSbs"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&Domicilio"
      TabPicture(1)   =   "PersonasMantenimiento.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPersDireccDomicilio"
      Tab(1).Control(1)=   "lblPersDireccCondicion"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "txtPersDireccDomicilio"
      Tab(1).Control(4)=   "cmbPersDireccCondicion"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "txtValComercial"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame2 
         Caption         =   "Teléfonos Fijos"
         Height          =   600
         Left            =   240
         TabIndex        =   124
         Top             =   960
         Width           =   4815
         Begin VB.TextBox txtPersTelefono 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   126
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtPersTelefono2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   125
            Top             =   200
            Width           =   1275
         End
         Begin VB.Label lblPersTelefono 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblPersTelefono2 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo:"
            Height          =   195
            Left            =   2640
            TabIndex        =   127
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Teléfonos Celulares"
         Height          =   615
         Left            =   240
         TabIndex        =   117
         Top             =   1560
         Width           =   6400
         Begin VB.TextBox txtCel1 
            Enabled         =   0   'False
            Height          =   300
            Left            =   840
            MaxLength       =   15
            TabIndex        =   120
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtCel2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   119
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtCel3 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5040
            MaxLength       =   15
            TabIndex        =   118
            Top             =   200
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Celular 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Celular 2:"
            Height          =   195
            Left            =   2160
            TabIndex        =   122
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Celular 3:"
            Height          =   195
            Left            =   4320
            TabIndex        =   121
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Correos Electrónicos"
         Height          =   615
         Left            =   240
         TabIndex        =   112
         Top             =   2160
         Width           =   7155
         Begin VB.TextBox TxtEmail 
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            MaxLength       =   45
            TabIndex        =   114
            Top             =   200
            Width           =   2865
         End
         Begin VB.TextBox TxtEmail2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            MaxLength       =   45
            TabIndex        =   113
            Top             =   200
            Width           =   2745
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Email 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Email 2:"
            Height          =   195
            Left            =   3720
            TabIndex        =   115
            Top             =   240
            Width           =   555
         End
      End
      Begin Sicmact.EditMoney txtValComercial 
         Height          =   255
         Left            =   -70080
         TabIndex        =   104
         Top             =   2580
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.TextBox TxtSbs 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5265
         MaxLength       =   9
         TabIndex        =   10
         Top             =   630
         Width           =   1515
      End
      Begin VB.ComboBox CmbRela 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3345
         Width           =   2235
      End
      Begin MSMask.MaskEdBox txtPersNacCreac 
         Height          =   300
         Left            =   2640
         TabIndex        =   9
         Top             =   645
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPersCiiu 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2880
         Width           =   5760
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ubicacion Geografica"
         Height          =   1665
         Left            =   -74865
         TabIndex        =   58
         Top             =   420
         Width           =   7365
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   525
            Width           =   1815
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   2235
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1140
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4695
            TabIndex        =   63
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2265
            TabIndex        =   62
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   61
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4680
            TabIndex        =   60
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2235
            TabIndex        =   59
            Top             =   900
            Width           =   600
         End
      End
      Begin VB.ComboBox cmbPersDireccCondicion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73740
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2550
         Width           =   1695
      End
      Begin VB.TextBox txtPersDireccDomicilio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         MaxLength       =   100
         TabIndex        =   19
         Top             =   2190
         Width           =   5200
      End
      Begin VB.ComboBox cmbPersEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3330
         Width           =   2520
      End
      Begin VB.Label Label13 
         Caption         =   "Valor Comercial U$"
         Height          =   240
         Left            =   -71580
         TabIndex        =   103
         Top             =   2625
         Width           =   1440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Codigo SBS :"
         Height          =   195
         Left            =   4275
         TabIndex        =   86
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Relac. Inst :"
         Height          =   195
         Left            =   3855
         TabIndex        =   85
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblPersDireccCondicion 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Left            =   -74820
         TabIndex        =   57
         Top             =   2625
         Width           =   705
      End
      Begin VB.Label lblPersDireccDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   56
         Top             =   2220
         Width           =   630
      End
      Begin VB.Label lblPersEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   285
         TabIndex        =   54
         Top             =   3330
         Width           =   495
      End
      Begin VB.Label lblPersCIIU 
         AutoSize        =   -1  'True
         Caption         =   "CIIU"
         Height          =   195
         Left            =   285
         TabIndex        =   49
         Top             =   2920
         Width           =   315
      End
      Begin VB.Label lblPersNac 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento/Creación"
         Height          =   195
         Left            =   300
         TabIndex        =   48
         Top             =   645
         Width           =   2220
      End
   End
   Begin VB.ComboBox cmbPersPersoneria 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3615
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   105
      Width           =   2835
   End
   Begin TabDlg.SSTab SSTabs 
      Height          =   2400
      Left            =   75
      TabIndex        =   41
      Top             =   570
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   4233
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Persona &Natural"
      TabPicture(0)   =   "PersonasMantenimiento.frx":0342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNombreAM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersNatEstCiv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersNatHijos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblApCasada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPersNombreN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPeso"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblTpoSangre"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPersNatSexo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPersNombreAP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblTalla"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbPersNatEstCiv"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtPersNatHijos"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtPeso"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbPersNatSexo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtTalla"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtPersNombreAP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPersNombreAM"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtApellidoCasada"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtPersNombreN"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CboTipoSangre"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbNacionalidad"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkResidente"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Persona &Juridica"
      TabPicture(1)   =   "PersonasMantenimiento.frx":035E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPersNombre"
      Tab(1).Control(1)=   "lblPersJurSiglas"
      Tab(1).Control(2)=   "lblPersJurTpo"
      Tab(1).Control(3)=   "lblPersJurMagnitud"
      Tab(1).Control(4)=   "lblPersJurEmpleados"
      Tab(1).Control(5)=   "txtPersNombreRS"
      Tab(1).Control(6)=   "TxtSiglas"
      Tab(1).Control(7)=   "cmbPersJurTpo"
      Tab(1).Control(8)=   "cmbPersJurMagnitud"
      Tab(1).Control(9)=   "txtPersJurEmpleados"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Relación con Personas"
      TabPicture(2)   =   "PersonasMantenimiento.frx":037A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPersRelacNew"
      Tab(2).Control(1)=   "cmdPersRelacDel"
      Tab(2).Control(2)=   "FERelPers"
      Tab(2).Control(3)=   "cmdPersRelacEditar"
      Tab(2).Control(4)=   "cmdPersRelacCancelar"
      Tab(2).Control(5)=   "cmdPersRelacAceptar"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Fuente Ingreso"
      TabPicture(3)   =   "PersonasMantenimiento.frx":0396
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FEFteIng"
      Tab(3).Control(1)=   "CmdFteIngEditar"
      Tab(3).Control(2)=   "CmdFteIngEliminar"
      Tab(3).Control(3)=   "CmdFteIngNuevo"
      Tab(3).Control(4)=   "CmdPersFteConsultar"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Fir&ma"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CmdActFirma"
      Tab(4).Control(1)=   "IDBFirma"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Ref Comercial"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "feRefComercial"
      Tab(5).Control(1)=   "cmdRefComAcepta"
      Tab(5).Control(2)=   "cmdRefComCancela"
      Tab(5).Control(3)=   "cmdRefComEdita"
      Tab(5).Control(4)=   "cmdRefComElimina"
      Tab(5).Control(5)=   "cmdRefComNuevo"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Ref Bancaria"
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdRefBanAcepta"
      Tab(6).Control(1)=   "cmdRefBanCancela"
      Tab(6).Control(2)=   "cmdRefBanEdita"
      Tab(6).Control(3)=   "cmdRefBanElimina"
      Tab(6).Control(4)=   "cmdRefBanNuevo"
      Tab(6).Control(5)=   "feRefBancaria"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Pat. Vehicular"
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdPatVehNuevo"
      Tab(7).Control(1)=   "cmdPatVehElimina"
      Tab(7).Control(2)=   "cmdPatVehEdita"
      Tab(7).Control(3)=   "cmdPatVehCancela"
      Tab(7).Control(4)=   "cmdPatVehAcepta"
      Tab(7).Control(5)=   "fePatVehicular"
      Tab(7).ControlCount=   6
      Begin PryImage.ImageDB IDBFirma 
         Height          =   1905
         Left            =   -73020
         TabIndex        =   111
         Top             =   330
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   3360
         Enabled         =   0   'False
      End
      Begin Sicmact.FlexEdit fePatVehicular 
         Height          =   1860
         Left            =   -74910
         TabIndex        =   110
         Top             =   420
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   3281
         Cols0           =   6
         HighLight       =   1
         EncabezadosNombres=   "#-Marca-Año Fabrica-Valor Comercial U$-Condicion-C"
         EncabezadosAnchos=   "400-3300-1200-1700-2200-0"
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
         ColumnasAEditar =   "X-1-2-3-4-X"
         ListaControles  =   "0-0-0-0-3-0"
         EncabezadosAlineacion=   "C-L-R-R-L-C"
         FormatosEdit    =   "0-0-3-2-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdPatVehAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -65925
         TabIndex        =   109
         Top             =   1665
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -65925
         TabIndex        =   108
         Top             =   1965
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65940
         TabIndex        =   107
         Top             =   810
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -65940
         TabIndex        =   106
         Top             =   1125
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65940
         TabIndex        =   105
         Top             =   495
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefBanAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66000
         TabIndex        =   101
         Top             =   1635
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66000
         TabIndex        =   100
         Top             =   1980
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   99
         Top             =   795
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   98
         Top             =   1110
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   97
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefComNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   96
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   95
         Top             =   1125
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   93
         Top             =   810
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -65970
         TabIndex        =   92
         Top             =   1950
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -65970
         TabIndex        =   91
         Top             =   1650
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkResidente 
         Alignment       =   1  'Right Justify
         Caption         =   "Residente"
         Enabled         =   0   'False
         Height          =   270
         Left            =   7815
         TabIndex        =   90
         Top             =   645
         Width           =   1770
      End
      Begin VB.ComboBox cmbNacionalidad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5865
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   1875
         Width           =   1560
      End
      Begin VB.ComboBox CboTipoSangre 
         Height          =   315
         Left            =   8730
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   1470
         Width           =   1125
      End
      Begin VB.TextBox txtPersNombreN 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1545
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1890
         Width           =   2895
      End
      Begin VB.TextBox TxtApellidoCasada 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1545
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1515
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAM 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1545
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAP 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   0
         Top             =   675
         Width           =   2895
      End
      Begin VB.CommandButton CmdPersFteConsultar 
         Caption         =   "Consultar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   81
         Top             =   1875
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66150
         TabIndex        =   78
         Top             =   1635
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66150
         TabIndex        =   77
         Top             =   1965
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   75
         Top             =   1110
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   74
         Top             =   795
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   73
         Top             =   810
         Width           =   1140
      End
      Begin VB.TextBox TxtTalla 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8730
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1005
         Width           =   870
      End
      Begin VB.ComboBox cmbPersNatSexo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   1545
      End
      Begin VB.TextBox TxtPeso 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8730
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1905
         Width           =   870
      End
      Begin Sicmact.FlexEdit FEFteIng 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   34
         Top             =   435
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3149
         Cols0           =   6
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Tipo-Razon Social-Moneda-CodPersFI-Indice"
         EncabezadosAnchos=   "400-1200-5000-1300-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-L-L-L"
         FormatosEdit    =   "3-1-1-1-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit FERelPers 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   31
         Top             =   435
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   3149
         Cols0           =   7
         FixedCols       =   0
         HighLight       =   2
         EncabezadosNombres=   "Item-Codigo-Nombres-Relacion-Beneficiario-Porcentaje-Asist.Med.Priv."
         EncabezadosAnchos=   "400-1300-4000-1400-2000-1200-2500"
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
         ColumnasAEditar =   "X-1-X-3-4-5-6"
         ListaControles  =   "0-1-0-3-3-0-3"
         EncabezadosAlineacion=   "C-L-L-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         CantEntero      =   3
         TextArray0      =   "Item"
         lbFlexDuplicados=   0   'False
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.CommandButton CmdActFirma 
         Caption         =   "&Actualizar Firma"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67740
         TabIndex        =   35
         Top             =   540
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdPersRelacDel 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   33
         Top             =   1125
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacNew 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   32
         Top             =   495
         Width           =   1140
      End
      Begin VB.TextBox txtPersJurEmpleados 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1800
         Width           =   1380
      End
      Begin VB.ComboBox cmbPersJurMagnitud 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1305
         Width           =   3000
      End
      Begin VB.ComboBox cmbPersJurTpo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1320
         Width           =   3000
      End
      Begin VB.TextBox TxtSiglas 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         MaxLength       =   15
         TabIndex        =   39
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox txtPersNombreRS 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73680
         MaxLength       =   150
         TabIndex        =   27
         Top             =   600
         Width           =   7815
      End
      Begin VB.TextBox txtPersNatHijos 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5865
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   1470
         Width           =   300
      End
      Begin VB.ComboBox cmbPersNatEstCiv 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1065
         Width           =   1560
      End
      Begin Sicmact.FlexEdit feRefComercial 
         Height          =   1785
         Left            =   -74940
         TabIndex        =   94
         Top             =   480
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   3149
         Cols0           =   5
         HighLight       =   2
         EncabezadosNombres=   "#-Nombre/Razón Social-Relacion-Telefono-C"
         EncabezadosAnchos=   "350-4900-1500-2000-0"
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
         ColumnasAEditar =   "X-1-2-3-X"
         ListaControles  =   "0-0-3-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         CantEntero      =   3
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit feRefBancaria 
         Height          =   1785
         Left            =   -74925
         TabIndex        =   102
         Top             =   510
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   3149
         Cols0           =   7
         HighLight       =   2
         EncabezadosNombres=   "#-Codigo-Referencia Bancaria-Nro Cuenta-Nro Tarjeta-Linea Cred U$-Item"
         EncabezadosAnchos=   "350-1500-3500-1800-1800-1500-0"
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
         ColumnasAEditar =   "X-1-X-3-4-5-X"
         ListaControles  =   "0-1-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label12 
         Caption         =   "Nacionalidad"
         Height          =   240
         Left            =   4815
         TabIndex        =   88
         Top             =   1965
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "m."
         Height          =   195
         Left            =   9690
         TabIndex        =   83
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kg."
         Height          =   195
         Left            =   9705
         TabIndex        =   82
         Top             =   1965
         Width           =   240
      End
      Begin VB.Label LblTalla 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
         Height          =   195
         Left            =   7830
         TabIndex        =   72
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label lblPersNombreAP 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   668
         Width           =   1155
      End
      Begin VB.Label lblPersNatSexo 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Left            =   4815
         TabIndex        =   70
         Top             =   675
         Width           =   375
      End
      Begin VB.Label LblTpoSangre 
         AutoSize        =   -1  'True
         Caption         =   "T. Sangre"
         Height          =   195
         Left            =   7830
         TabIndex        =   69
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label lblPeso 
         AutoSize        =   -1  'True
         Caption         =   "Peso "
         Height          =   195
         Left            =   7875
         TabIndex        =   68
         Top             =   1965
         Width           =   405
      End
      Begin VB.Label lblPersNombreN 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Left            =   135
         TabIndex        =   67
         Top             =   1913
         Width           =   630
      End
      Begin VB.Label lblApCasada 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Casada"
         Height          =   195
         Left            =   135
         TabIndex        =   64
         Top             =   1538
         Width           =   1140
      End
      Begin VB.Label lblPersJurEmpleados 
         AutoSize        =   -1  'True
         Caption         =   "N° Empleados"
         Height          =   195
         Left            =   -74880
         TabIndex        =   55
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblPersJurMagnitud 
         AutoSize        =   -1  'True
         Caption         =   "Magnitud Empresarial"
         Height          =   195
         Left            =   -70560
         TabIndex        =   53
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label lblPersJurTpo 
         AutoSize        =   -1  'True
         Caption         =   "Tpo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   52
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label lblPersJurSiglas 
         AutoSize        =   -1  'True
         Caption         =   "Siglas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   51
         Top             =   960
         Width           =   420
      End
      Begin VB.Label lblPersNombre 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social"
         Height          =   195
         Left            =   -74880
         TabIndex        =   50
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblPersNatHijos 
         AutoSize        =   -1  'True
         Caption         =   "N°Hijos"
         Height          =   195
         Left            =   4800
         TabIndex        =   46
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label lblPersNatEstCiv 
         AutoSize        =   -1  'True
         Caption         =   "Est. Civil"
         Height          =   195
         Left            =   4815
         TabIndex        =   45
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label lblPersNombreAM 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   1125
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab SSTIdent 
      Height          =   2940
      Left            =   7965
      TabIndex        =   43
      Top             =   3195
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5186
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Identificación"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPersIDCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPersIDAceptar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPersIDedit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPersIDDel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FEDocs"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPersIDnew"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdPersIDnew 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   22
         Top             =   2460
         Width           =   885
      End
      Begin Sicmact.FlexEdit FEDocs 
         Height          =   1920
         Left            =   135
         TabIndex        =   21
         Top             =   435
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   3387
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Tipo-Numero"
         EncabezadosAnchos=   "350-1200-1200"
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
         ColumnasAEditar =   "X-1-2"
         ListaControles  =   "0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         AvanceCeldas    =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdPersIDDel 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         TabIndex        =   24
         Top             =   2475
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDedit 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1125
         TabIndex        =   23
         Top             =   2460
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   180
         TabIndex        =   79
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   1125
         TabIndex        =   80
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgImg 
      Height          =   480
      Left            =   10695
      TabIndex        =   84
      Top             =   2475
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   847
      Filtro          =   "Archivos bmp|*.bmp|Archivos jpg|*.jpg|Archivos gif|*.gif|Todos los Archivos|*.*"
      Altura          =   280
   End
   Begin VB.Label lblPersPersoneria 
      AutoSize        =   -1  'True
      Caption         =   "Personería:"
      Height          =   195
      Left            =   2745
      TabIndex        =   42
      Top             =   165
      Width           =   825
   End
   Begin VB.Label lblPersCod 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   60
      TabIndex        =   40
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "frmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************
Private Enum TTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Enum TPersonaTipoInicio
    PersonaConsulta = 1
    PersonaActualiza = 2
End Enum

Dim NomMoverSSTabs As Integer

Dim oBuscaPersona As UPersona
Dim oPersona As DPersona
Dim Nivel1() As String
Dim ContNiv1 As Long
Dim Nivel2() As String
Dim ContNiv2 As Long
Dim Nivel3() As String
Dim ContNiv3 As Long
Dim Nivel4() As String
Dim ContNiv4 As Long
Dim Nivel5() As String
Dim ContNiv5 As Long
Dim bEstadoCargando As Boolean

'Para flexEdit de Relacion con Personas
Dim cmdPersRelaEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FERelPersNoMoverdeFila As Integer

'Para flexEdit de Documentos
Dim cmdPersDocEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FEDocsPersNoMoverdeFila As Integer

'Para Fuentes de Ingreso
Dim cmdPersFteIngresoEjecutado As Integer
Dim FEFtePersNoMoverdeFila As Integer

'Para Ref Comercial
Dim cmdPersRefComercialEjecutado As Integer
Dim FERefComPersNoMoverdeFila As Integer
Dim lnNumRefCom As Integer

'Para Ref Bancaria
Dim cmdPersRefBancariaEjecutado As Integer
Dim FERefBanPersNoMoverdeFila As Integer

'Para PAt Vehicular
Dim cmdPersPatVehicularEjecutado As Integer
Dim FEPatVehPersNoMoverdeFila As Integer
Dim lnNumPatVeh As Integer

Dim BotonEditar As Boolean
Dim BotonNuevo As Boolean

'Para Persona Nueva desde BuscaPersona
Dim sPersCodNombre As String
Dim bBuscaNuevo As Boolean
Dim bNuevaPersona As Boolean
Dim bPersonaAct As Boolean
Dim nPos As Integer
Dim bPermisoEditarTodo As Boolean 'EJVG20111219
'WIOR 20130827 *************************************
Private fbPermisoCargo As Boolean
Private fsNombreActual As String
Private rsDocPersActual As ADODB.Recordset
Private rsDocPersUlt As ADODB.Recordset
'WIOR FIN ******************************************

Public Sub Registrar()
    Me.Caption = "Personas:Registro"
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = True
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Caption = "Persona : Registro"
    bPermisoEditarTodo = True
    Me.Show 1
End Sub

Public Sub Mantenimeinto()
    Me.Caption = "Personas:Mantenimiento"
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = False
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Caption = "Persona : Mantenimiento"
    Me.Show 1
End Sub

Public Sub Consultar()
    Me.Caption = "Personas:Consulta"
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Show 1
End Sub

Public Function PersonaNueva() As String
    On Error GoTo LblErrorNew
    sPersCodNombre = ""
    bBuscaNuevo = True
    Me.Show 1
    Call CmdNuevo_Click
    PersonaNueva = sPersCodNombre
    Exit Function
LblErrorNew:
    Select Case Err.Number
        Case 400
            MsgBox "No se Puede Crear una Persona desde el Mantenimiento de Persona", vbInformation, "Aviso"
        Case Else
            MsgBox Err.Description, vbInformation, "Aviso"
    End Select
    
End Function

Public Sub Inicio(ByVal cPersCod As String, ByVal pTipoInicio As TPersonaTipoInicio)
    
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New DPersona
    TxtBCodPers.Text = cPersCod
    Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
    If oPersona.PersCodigo = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    Call CargaDatos
    TxtBCodPers.Enabled = False
    cmdNuevo.Enabled = False
    
    If pTipoInicio = PersonaActualiza Then
        cmdEditar.Enabled = True
    Else
        cmdEditar.Enabled = False
    End If
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
        
    
    Me.Show 1
End Sub

Private Function ValidaDatosDocumentos() As Boolean
Dim I, J As Integer
Dim bEnc As Boolean
    ValidaDatosDocumentos = True
    
    'Verifica el Tipo de Documento
    If Len(Trim(FEDocs.TextMatrix(FEDocs.row, 1))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Documento", vbInformation, "Aviso"
        If FEDocs.Enabled And SSTIdent.Enabled Then FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    'Verifica el Numero de Documento
    If Len(Trim(FEDocs.TextMatrix(FEDocs.row, 2))) = 0 Then
        MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "Aviso"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    'Verifica Duplicidad de Documento
    bEnc = False
    For I = 1 To FEDocs.Rows - 2
        For J = I + 1 To FEDocs.Rows - 1
            If Trim(Right(FEDocs.TextMatrix(I, 1), 20)) = Trim(Right(FEDocs.TextMatrix(J, 1), 20)) Then
                bEnc = True
                Exit For
            End If
        Next J
        If bEnc Then
            Exit For
        End If
    Next I
    If bEnc Then
        MsgBox "Existe un Documento Duplicado", vbInformation, "Aviso"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    '******************************************************************
    'Verfica Tipo de Valores de Documento
    For I = 1 To FEDocs.Rows - 1
        If Trim(FEDocs.TextMatrix(I, 1)) <> "" Then
            If CInt(Right(Trim(FEDocs.TextMatrix(I, 1)), 2)) = gPersIdDNI Then
                For J = 1 To Len(Trim(FEDocs.TextMatrix(I, 2)))
                    If (Mid(FEDocs.TextMatrix(I, 2), J, 1) < "0" Or Mid(FEDocs.TextMatrix(I, 2), J, 1) > "9") Then
                       MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "Aviso"
                       FEDocs.SetFocus
                       ValidaDatosDocumentos = False
                       Exit Function
                    End If
                Next J
            End If
            If CInt(Right(Trim(FEDocs.TextMatrix(I, 1)), 2)) = gPersIdRUC Then
                 For J = 1 To Len(Trim(FEDocs.TextMatrix(I, 2)))
                    If (Mid(FEDocs.TextMatrix(I, 2), J, 1) < "0" Or Mid(FEDocs.TextMatrix(I, 2), J, 1) > "9") Then
                       MsgBox "Uno de los Digitos del RUC no es un Numero", vbInformation, "Aviso"
                       FEDocs.SetFocus
                       ValidaDatosDocumentos = False
                       Exit Function
                    End If
                Next J
            End If
        End If
    Next I
    
    'Verfica Longitud de Documento
    '******************************************************************
    For I = 1 To FEDocs.Rows - 1
        If Trim(FEDocs.TextMatrix(I, 1)) <> "" Then
            If CInt(Right(Trim(FEDocs.TextMatrix(I, 1)), 2)) = gPersIdDNI Then
                If Len(Trim(FEDocs.TextMatrix(I, 2))) <> gnNroDigitosDNI Then
                    MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    FEDocs.SetFocus
                    ValidaDatosDocumentos = False
                    Exit Function
                End If
            End If
            If CInt(Right(Trim(FEDocs.TextMatrix(I, 1)), 2)) = gPersIdRUC Then
                If Len(Trim(FEDocs.TextMatrix(I, 2))) <> gnNroDigitosRUC Then
                    MsgBox "RUC No es de " & gnNroDigitosRUC & " digitos", vbInformation, "Aviso"
                    FEDocs.SetFocus
                    ValidaDatosDocumentos = False
                    Exit Function
                End If
            End If
        End If
    Next I
    
End Function

Private Function ValidaDatosPersRelacion() As Boolean
Dim I As Integer

    ValidaDatosPersRelacion = True
    
    'Valida Titular No Este como Relacion
    For I = 1 To Me.FERelPers.Rows - 1
        If FERelPers.TextMatrix(I, 1) = TxtBCodPers.Text Then
            MsgBox "No se puede Agregar al Titular en la Relacion de Personas", vbInformation, "Aviso"
            ValidaDatosPersRelacion = False
            FERelPers.SetFocus
            Exit Function
        End If
    Next I
    
    'Falta Persona a Relacionar
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 1))) = 0 Then
        MsgBox "Falta Ingresar la Persona con la que se va a Relacionar", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta Ingresar el tipo de Relacion
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 3))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Relacion con la Persona", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta el Tipo de Beneficio
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 4))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Beneficio con la Persona", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta ingresar el tipo de Asistenacia Medica Privada AMP
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 6))) = 0 Then
        MsgBox "Falta Ingresar el Tipo Asist. Med. Privada", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatosRefComercial() As Boolean
    
    ValidaDatosRefComercial = True

    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 1))) = 0 Then 'Nombre/Razon Social
        MsgBox "Falta Ingresar Nombre o Razón Social de la Referencia Comercial", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 2))) = 0 Then 'Tipo de Referencia
        MsgBox "Falta Ingresar el Tipo de Referencia Comercial", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 3))) = 0 Then 'Número Telefónico
        MsgBox "Falta Ingresar Número Telefónico de Referencia Comercial", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatosRefBancaria() As Boolean
    
    ValidaDatosRefBancaria = True

    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 1))) = 0 Then '
        MsgBox "Falta Seleccionar la Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 3))) = 0 Then 'Número de Cuenta
        MsgBox "Falta Ingresar el Número de Cuenta de la Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 4))) = 0 Then 'Número de Tarjeta
        MsgBox "Falta Ingresar Número de Tarjeta de Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 5))) = 0 Then 'Línea de Crédito
        MsgBox "Falta Ingresar Línea de Crédito de Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(feRefBancaria.TextMatrix(feRefBancaria.row, 3)) > 20 Then
        MsgBox "Longitud de Cuenta no válida", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(feRefBancaria.TextMatrix(feRefBancaria.row, 4)) > 20 Then
        MsgBox "Longitud de Numero de Tarjeta no válida", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatosPatVehicular() As Boolean
    
    ValidaDatosPatVehicular = True

    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 1))) = 0 Then 'Marca
        MsgBox "Falta Ingresar la Marca del Vehículo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 2))) = 0 Then 'Fecha de Fabricacion
        MsgBox "Falta Ingresar la Fecha de Fabricación", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    ElseIf CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) > Year(gdFecSis) Then
        MsgBox "Año de Fabricación no puede ser mayor que Año actual", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 3))) = 0 Then 'Valor Comercial
        MsgBox "Falta Ingresar el Valor Comercial del Vehículo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 4))) = 0 Then 'Condición del Patrimomio
        MsgBox "Falta Seleccionar la Condición del Patrimonio Vehicular", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
End Function

Private Function ValidaControles() As Boolean
    Dim CadTmp As String
    
    ValidaControles = True
    
    If cmdPersIDAceptar.Visible Then
            MsgBox "Pulse Aceptar para Confirmar Documento", vbInformation, "Aviso"
            cmdPersIDAceptar.SetFocus
            ValidaControles = False
            Exit Function
    End If
           
    If gnPersPersoneria = 1 And gsCodArea <> "022" Then '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
    
        If cmbPersPersoneria.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Personeria", vbInformation, "Aviso"
            cmbPersPersoneria.SetFocus
            ValidaControles = False
            Exit Function
        End If
        
    End If
    
    If oPersona.Personeria = gPersonaNat Then
    'Valida Controles de Persona Natural
        If Len(Trim(txtPersNombreAP.Text)) = 0 Then
            MsgBox "Falta Ingresar el Apellido Paterno", vbInformation, "Aviso"
            'txtPersNombreAP.SetFocus
            If txtPersNombreAP.Enabled Then 'EJVG20120120
                txtPersNombreAP.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNombreAM.Text)) = 0 Then
            MsgBox "Falta Ingresar el Apellido Materno", vbInformation, "Aviso"
            'txtPersNombreAM.SetFocus
            If txtPersNombreAM.Enabled Then 'EJVG20120120
                txtPersNombreAM.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNombreN.Text)) = 0 Then
            MsgBox "Falta Ingresar Nombres de la Persona", vbInformation, "Aviso"
            'txtPersNombreN.SetFocus
            If txtPersNombreN.Enabled Then 'EJVG20120120
                txtPersNombreN.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        
        If gnPersPersoneria = 1 And gsCodArea <> "022" Then '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
        
            If cmbPersNatSexo.ListIndex = -1 Then
                MsgBox "Falta Seleccionar el Sexo de la Persona", vbInformation, "Aviso"
                'cmbPersNatSexo.SetFocus
                If cmbPersNatSexo.Enabled Then 'EJVG20120120
                    cmbPersNatSexo.SetFocus
                End If
                ValidaControles = False
                SSTabs.Tab = 0
                Exit Function
            End If
            If cmbPersNatEstCiv.ListIndex = -1 Then
                MsgBox "Falta Seleccionar el Estado Civil de la Persona", vbInformation, "Aviso"
                'cmbPersNatEstCiv.SetFocus
                If cmbPersNatEstCiv.Enabled Then 'EJVG20120120
                    cmbPersNatEstCiv.SetFocus
                End If
                ValidaControles = False
                SSTabs.Tab = 0
                Exit Function
            End If
            If Len(Trim(txtPersNatHijos.Text)) = 0 Then
                MsgBox "Falta Ingresar el Numero de Hijos", vbInformation, "Aviso"
                txtPersNatHijos.SetFocus
                ValidaControles = False
                SSTabs.Tab = 0
                Exit Function
            End If
            
        End If
'        If Len(Trim(TxtTipoSangre.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Tipo de Sangre de la Persona de la Persona", vbInformation, "Aviso"
'            ValidaControles = False
'            Exit Function
'        End If
    Else
    'Valida Controles de Persona Juridica
        If Len(Trim(txtPersNombreRS.Text)) = 0 Then
            MsgBox "Falta Ingresar la razon Social de la Persona", vbInformation, "Aviso"
            txtPersNombreRS.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
    
        If Len(Trim(TxtSiglas.Text)) = 0 Then
            MsgBox "Falta Ingresar la Siglas de la Persona", vbInformation, "Aviso"
            TxtSiglas.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If cmbPersJurTpo.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Tipo de Persona Juridica", vbInformation, "Aviso"
            cmbPersJurTpo.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If cmbPersJurMagnitud.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Magnitud Empresarial de la Persona", vbInformation, "Aviso"
            cmbPersJurMagnitud.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If Len(Trim(txtPersJurEmpleados.Text)) = 0 Then
            MsgBox "Falta Ingresar el Numero de Empleados ", vbInformation, "Aviso"
            txtPersJurEmpleados.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
    End If
    
    
    'Valida Datos Generales
    If gnPersPersoneria = 1 And gsCodArea <> "022" Then '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
    
        CadTmp = ValidaFecha(txtPersNacCreac.Text)
        If Len(CadTmp) > 0 Then
            MsgBox CadTmp, vbInformation, "Aviso"
            'txtPersNacCreac.SetFocus
            If txtPersNacCreac.Enabled Then 'EJVG20120120
                txtPersNacCreac.SetFocus
            End If
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    
    
        If CboPersCiiu.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Ciiu de la Persona", vbInformation, "Aviso"
            CboPersCiiu.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    
        If cmbPersEstado.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Estado de la Persona", vbInformation, "Aviso"
            cmbPersEstado.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    
        If cmbPersUbiGeo(4).ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Ubicacion Geografica de la Persona", vbInformation, "Aviso"
            cmbPersUbiGeo(4).SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 1
            Exit Function
        End If
    
        If Len(Trim(txtPersDireccDomicilio.Text)) = 0 Then
            MsgBox "Falta Ingresar el Domicilio de la Persona", vbInformation, "Aviso"
            txtPersDireccDomicilio.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 1
            Exit Function
        End If
    
        If cmbPersDireccCondicion.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Condicion del Domicilio de la Persona", vbInformation, "Aviso"
            cmbPersDireccCondicion.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 1
            Exit Function
        End If
    End If
    
    If Trim(FEDocs.TextMatrix(1, 0)) = "" And Trim(FEDocs.TextMatrix(1, 1)) = "" Then
        MsgBox "Falta Ingresar algun Documento de Identidad", vbInformation, "Aviso"
        SSTIdent.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If Trim(txtPersNatHijos.Text) = "" Then txtPersNatHijos.Text = "0"
    
    If CInt(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
    
        If gnPersPersoneria = 1 And gsCodArea <> "022" Then '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
    
            If CInt(txtPersNatHijos.Text) > 15 Then
                MsgBox "Numero de Hijos Incorrecto", vbInformation, "Aviso"
                txtPersNatHijos.SetFocus
                ValidaControles = False
                Exit Function
            End If
        
        End If
    
    End If
    
    If CInt(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
    
        If gnPersPersoneria = 1 And gsCodArea <> "022" Then '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
        
            If cmbNacionalidad.ListIndex = -1 Then
                MsgBox "Falta Seleccionar la Nacionalidad de la Persona", vbInformation, "Aviso"
                If cmbNacionalidad.Enabled Then cmbNacionalidad.SetFocus
                ValidaControles = False
                SSTDatosGen.Tab = 0
                Exit Function
            End If
        
            If CmbRela.ListIndex = -1 Then
                MsgBox "Falta Seleccionar la Relacion de la Persona con la Institución", vbInformation, "Aviso"
                CmbRela.SetFocus
                ValidaControles = False
                SSTDatosGen.Tab = 0
                Exit Function
            End If
            
        End If
        
    End If
        
    'EJVG20111207 *********************Obliga un Tel. Fijo o Celular, verifica los Emails
    If Len(Trim(txtPersTelefono.Text)) = 0 And Len(Trim(txtPersTelefono2.Text)) = 0 And Len(Trim(txtCel1.Text)) = 0 And Len(Trim(txtCel2.Text)) = 0 And Len(Trim(txtCel3.Text)) = 0 Then
        MsgBox "Falta Ingresar un Teléfono Fijo o Celular", vbInformation, "Aviso"
        txtPersTelefono.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    If TxtEmail.Text <> "" Then
        If EsEmailValido(TxtEmail.Text) = False Then
            MsgBox "Ingrese un Email válido", vbInformation, "Aviso"
            TxtEmail.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    End If
    If TxtEmail2.Text <> "" Then
        If EsEmailValido(TxtEmail2.Text) = False Then
            MsgBox "Ingrese un Email válido", vbInformation, "Aviso"
            TxtEmail2.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    End If
End Function

Private Sub HabilitaControlesPersona(ByVal pbBloqueo As Boolean)
    
    cmbPersPersoneria.Enabled = pbBloqueo
    
    If Not oPersona Is Nothing Then
        If oPersona.Personeria = gPersonaNat Then
            'Ficha de Persona natural
            Call HabilitaFichaPersonaNat(pbBloqueo)
        Else
            'Ficha de Persona Juridica
            Call HabilitaFichaPersonaJur(pbBloqueo)
        End If
    Else
        'Ficha de Persona natural
        Call HabilitaFichaPersonaNat(pbBloqueo)
        'Ficha de Persona Juridica
        Call HabilitaFichaPersonaJur(Not pbBloqueo)
    End If
            
    TxtBCodPers.Enabled = Not pbBloqueo
    'Ficha de Relaciones de Persona
    FERelPers.lbEditarFlex = pbBloqueo
    cmdPersRelacNew.Enabled = pbBloqueo
    cmdPersRelacEditar.Enabled = pbBloqueo
    cmdPersRelacDel.Enabled = pbBloqueo
    cmdPersRelacAceptar.Enabled = pbBloqueo
    cmdPersRelacCancelar.Enabled = pbBloqueo
    
    'Ficha RefComercial
    feRefComercial.lbEditarFlex = pbBloqueo
    cmdRefComNuevo.Enabled = pbBloqueo
    cmdRefComEdita.Enabled = pbBloqueo
    cmdRefComElimina.Enabled = pbBloqueo
    cmdRefComAcepta.Enabled = pbBloqueo
    cmdRefComCancela.Enabled = pbBloqueo
    
    'Ficha Referencia Bancaria
    feRefBancaria.lbEditarFlex = pbBloqueo
    cmdRefBanNuevo.Enabled = pbBloqueo
    cmdRefBanEdita.Enabled = pbBloqueo
    cmdRefBanElimina.Enabled = pbBloqueo
    cmdRefBanAcepta.Enabled = pbBloqueo
    cmdRefBanCancela.Enabled = pbBloqueo
    
    'Ficha Patrimonio Vehicular
    fePatVehicular.lbEditarFlex = pbBloqueo
    cmdPatVehNuevo.Enabled = pbBloqueo
    cmdPatVehEdita.Enabled = pbBloqueo
    cmdPatVehElimina.Enabled = pbBloqueo
    cmdPatVehAcepta.Enabled = pbBloqueo
    cmdPatVehCancela.Enabled = pbBloqueo
    
    'Ficha de Fuentes de Ingreso
    Call HabilitaControlesPersonaFtesIngreso(pbBloqueo)
    'Firma
    CmdActFirma.Enabled = pbBloqueo
        
    'Ficha de Datos Generales
    txtPersNacCreac.Enabled = pbBloqueo
    txtPersTelefono.Enabled = pbBloqueo
    txtPersTelefono2.Enabled = pbBloqueo
    'EJVG20111209 ***********************
    txtCel1.Enabled = pbBloqueo
    txtCel2.Enabled = pbBloqueo
    txtCel3.Enabled = pbBloqueo
        
    TxtEmail.Enabled = pbBloqueo
    TxtEmail2.Enabled = pbBloqueo 'EJVG20111209
    CboPersCiiu.Enabled = pbBloqueo
    cmbPersEstado.Enabled = pbBloqueo
    
    'Ficha de Ubicacion Geografica
    cmbPersUbiGeo(0).Enabled = pbBloqueo
    cmbPersUbiGeo(1).Enabled = pbBloqueo
    cmbPersUbiGeo(2).Enabled = pbBloqueo
    cmbPersUbiGeo(3).Enabled = pbBloqueo
    cmbPersUbiGeo(4).Enabled = pbBloqueo
    txtPersDireccDomicilio.Enabled = pbBloqueo
    cmbPersDireccCondicion.Enabled = pbBloqueo
    txtValComercial.Enabled = pbBloqueo
    TxtSbs.Enabled = pbBloqueo
    CmbRela.Enabled = pbBloqueo
    'Ficha de Identificacion
    SSTIdent.Enabled = pbBloqueo
    FEDocs.lbEditarFlex = IIf(pbBloqueo, False, True)
    cmdPersIDnew.Enabled = pbBloqueo
    cmdPersIDedit.Enabled = pbBloqueo
    cmdPersIDDel.Enabled = pbBloqueo
    CmdPersAceptar.Enabled = True
    
End Sub
Private Sub HabilitaControlesPersonaFtesIngreso(ByVal pbBloqueo As Boolean)
    FEFteIng.lbEditarFlex = pbBloqueo
    CmdFteIngNuevo.Enabled = pbBloqueo
    CmdFteIngEditar.Enabled = pbBloqueo
    CmdFteIngEliminar.Enabled = pbBloqueo
    CmdPersFteConsultar.Enabled = pbBloqueo
End Sub
Private Sub HabilitaFichaPersonaNat(ByVal pbFicActiva As Boolean)

    lblPersNombreAP.Enabled = pbFicActiva
    txtPersNombreAP.Enabled = pbFicActiva
    lblPersNombreAM.Enabled = pbFicActiva
    txtPersNombreAM.Enabled = pbFicActiva
    lblApCasada.Enabled = pbFicActiva
    TxtApellidoCasada.Enabled = pbFicActiva
    lblPersNombreN.Enabled = pbFicActiva
    txtPersNombreN.Enabled = pbFicActiva
    lblPersNatSexo.Enabled = pbFicActiva
    cmbPersNatSexo.Enabled = pbFicActiva
    lblPersNatEstCiv.Enabled = pbFicActiva
    cmbPersNatEstCiv.Enabled = pbFicActiva
    lblPersNatHijos.Enabled = pbFicActiva
    txtPersNatHijos.Enabled = pbFicActiva
    lblPeso.Enabled = pbFicActiva
    TxtPeso.Enabled = pbFicActiva
    LblTalla.Enabled = pbFicActiva
    TxtTalla.Enabled = pbFicActiva
    LblTpoSangre.Enabled = pbFicActiva
    CboTipoSangre.Enabled = pbFicActiva
    cmbNacionalidad.Enabled = pbFicActiva
    chkResidente.Enabled = pbFicActiva

End Sub
Private Sub HabilitaFichaPersonaJur(ByVal pbFicActiva As Boolean)
    txtPersNombreRS.Enabled = pbFicActiva
    TxtSiglas.Enabled = pbFicActiva
    cmbPersJurTpo.Enabled = pbFicActiva
    txtPersJurEmpleados.Enabled = pbFicActiva
    cmbPersJurMagnitud.Enabled = pbFicActiva
    lblPersNombre.Enabled = pbFicActiva
    lblPersJurSiglas.Enabled = pbFicActiva
    lblPersJurTpo.Enabled = pbFicActiva
    lblPersJurEmpleados.Enabled = pbFicActiva
    lblPersJurMagnitud.Enabled = pbFicActiva
End Sub

Private Sub CargaControles()
Dim sSql As String
Dim Conn As DConecta
Dim R As ADODB.Recordset
Dim I As Integer
Dim oConstante As DConstante
Dim oCtasIF As NCajaCtaIF

On Error GoTo ERRORCargaControles
    bEstadoCargando = True
    cmdPersDocEjecutado = 0
    cmdPersRelaEjecutado = 0
    cmdPersFteIngresoEjecutado = 0
    NomMoverSSTabs = -1
    FERelPersNoMoverdeFila = -1
    FEDocsPersNoMoverdeFila = -1
    FEFtePersNoMoverdeFila = -1
    Set oConstante = New DConstante
    Set Conn = New DConecta
    Conn.AbreConexion
    
    txtPersNacCreac.Text = Format(gdFecSis, "dd/mm/yyyy")
    'Carga Combo de Flex de Relaciones de Persona
    FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
    
    'Carga Combo de Documentos
    FEDocs.CargaCombo oConstante.RecuperaConstantes(gPersIdTipo)
    
    'Carga Combo de Tipos de Referencia Comercial
    feRefComercial.CargaCombo oConstante.RecuperaConstantes(3028)
    
    'Carga Combo de Condicion de Patrimonio Vehicular
    fePatVehicular.CargaCombo oConstante.RecuperaConstantes(3029)
    
    Set oConstante = Nothing
    
    'Carga TextBuscar Instituciones Financieras
    Set oCtasIF = New NCajaCtaIF
    feRefBancaria.psRaiz = "BANCOS"
    feRefBancaria.rsTextBuscar = oCtasIF.GetInstFinancieras("0[123]")
    Set oCtasIF = Nothing
    
    'Carga Tipos de Sangre
    Call CargaComboConstante(gPersTpoSangre, CboTipoSangre)
    
    'Carga Condiciones de Domicilio
    cmbPersDireccCondicion.Clear
    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
    
    'Carga Personeria
    cmbPersPersoneria.Clear
    Call CargaComboConstante(gPersPersoneria, cmbPersPersoneria)
    
    'Carga Tipos de Sexo de Personas
    cmbPersNatSexo.AddItem "FEMENINO" & Space(50) & "F"
    cmbPersNatSexo.AddItem "MASCULINO" & Space(50) & "M"
    
    'Carga Magnitud Empresarial
    cmbPersJurMagnitud.Clear
    Call CargaComboConstante(gPersJurMagnitud, cmbPersJurMagnitud)
        
    'Carga Condicion de Domicilio
    cmbPersDireccCondicion.Clear
    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
    
    'Carga Estado Civil
    cmbPersNatEstCiv.Clear
    Call CargaComboConstante(gPersEstadoCivil, cmbPersNatEstCiv)
    
    'Carga Combo Relaciones Con La Institucion
    CmbRela.Clear
    Call CargaComboConstante(gPersRelacionInst, CmbRela)
    
    'Carga Ubicaciones Geograficas
    Call CargaUbicacionesGeograficas
    
    'Carga Ciiu
    If gsCodCMAC = "102" Then
        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
             & "  union all " _
             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
             & " order by nro, cCIIUdescripcion "
    Else
        sSql = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
    End If
    
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        CboPersCiiu.AddItem Trim(R!cCIIUdescripcion) & Space(100) & Trim(R!cCIIUcod)
        R.MoveNext
    Loop
    R.Close
    
    'CARGA TIPOS DE PERSONA JURIDICA
    cmbPersJurTpo.Clear
    sSql = "Select cPersJurTpoCod,cPersJurTpoDesc  from persjurtpo Order by cPersJurTpoDesc"
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        cmbPersJurTpo.AddItem Trim(R!cPersJurTpoDesc) & Space(100) & Trim(R!cPersJurTpoCod)
        R.MoveNext
    Loop
    
    R.Close
    
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    bEstadoCargando = True
    Exit Sub
    
    
ERRORCargaControles:
    MsgBox Err.Description, vbExclamation, "Aviso"
    
End Sub
Private Sub CargaControlEstadoPersona(ByVal pnTipoPers As Integer)
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

On Error GoTo ERRORCargaControles
    Set Conn = New DConecta
    Conn.AbreConexion
    cmbPersEstado.Clear
    
    'Carga Estados de la Persona
    sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod = " & Trim(Str(gPersEstado)) & " and nConsValor <> " & Trim(Str(gPersEstado))
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        If pnTipoPers = gPersonaNat Then
            If Len(Trim(R!nConsValor)) = 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(Str(R!nConsValor)), 2)
            End If
        Else
            If Len(Trim(R!nConsValor)) > 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(R!nConsValor), 2)
            End If
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Sub
    
ERRORCargaControles:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub DistribuyeApellidos(ByVal bApellCasada As Boolean)
    If bApellCasada = True Then
        lblApCasada.Visible = True
        TxtApellidoCasada.Visible = True
        lblPersNombreAP.Top = 660
        lblPersNombreAP.Left = 120
        txtPersNombreAP.Top = 660
        txtPersNombreAP.Left = 1680
        lblPersNombreAM.Top = 1065
        lblPersNombreAM.Left = 120
        txtPersNombreAM.Top = 1065
        txtPersNombreAM.Left = 1680
        lblApCasada.Top = 1485
        lblApCasada.Left = 135
        TxtApellidoCasada.Top = 1485
        TxtApellidoCasada.Left = 1680
        lblPersNombreN.Top = 1860
        lblPersNombreN.Left = 135
        txtPersNombreN.Top = 1875
        txtPersNombreN.Left = 1680
    Else
        lblPersNombreAP.Top = 660
        lblPersNombreAP.Left = 120
        txtPersNombreAP.Top = 660
        txtPersNombreAP.Left = 1680
        lblPersNombreAM.Top = 1140
        lblPersNombreAM.Left = 120
        txtPersNombreAM.Top = 1140
        txtPersNombreAM.Left = 1680
        lblApCasada.Visible = False
        TxtApellidoCasada.Visible = False
        lblPersNombreN.Top = 1620
        lblPersNombreN.Left = 120
        txtPersNombreN.Top = 1620
        txtPersNombreN.Left = 1680
    End If
End Sub
Private Sub CargaDocumentos()
Dim I As Integer
    Call LimpiaFlex(FEDocs)
    For I = 0 To oPersona.NumeroDocumentos - 1
        FEDocs.AdicionaFila
        'Columna de Tipo de Documento
        FEDocs.TextMatrix(I + 1, 1) = Trim(oPersona.ObtenerTipoDoc(I))
        'Columna de Numero de Documento
        FEDocs.TextMatrix(I + 1, 2) = oPersona.ObtenerNumeroDoc(I)
    Next I
End Sub
Private Sub CargaRelacionesPersonas()
Dim I As Integer
    FERelPers.lbEditarFlex = True
    Call LimpiaFlex(FERelPers)
    For I = 0 To oPersona.NumeroRelacPers - 1
        FERelPers.AdicionaFila
        'Codigo
        FERelPers.TextMatrix(I + 1, 1) = oPersona.ObtenerRelaPersCodigo(I)
        'Apellidos y Nombres
        FERelPers.TextMatrix(I + 1, 2) = oPersona.ObtenerRelaPersNombres(I)
        'Relacion
        FERelPers.TextMatrix(I + 1, 3) = oPersona.ObtenerRelaPersRelacion(I)
        'Beneficiario
        FERelPers.TextMatrix(I + 1, 4) = oPersona.ObtenerRelaPersBenef(I)
        'Beneficiario Porcentaje
        FERelPers.TextMatrix(I + 1, 5) = Format(oPersona.ObtenerRelaPersBenefPorc(I), "#0.00")
        'Asistencia medica Privada
        FERelPers.TextMatrix(I + 1, 6) = oPersona.ObtenerRelaPersAMP(I)
    Next I
    FERelPers.lbEditarFlex = False
End Sub

Private Sub CargaRefComerciales()
Dim I As Integer

    feRefComercial.lbEditarFlex = True
    Call LimpiaFlex(feRefComercial)
    For I = 0 To oPersona.NumeroRefComercial - 1
        feRefComercial.AdicionaFila
        feRefComercial.TextMatrix(I + 1, 1) = oPersona.ObtenerRefComNombre(I) 'Nombre/Razón Social
        feRefComercial.TextMatrix(I + 1, 2) = oPersona.ObtenerRefComRelacion(I) 'Tipo de Relacion con la Referencia
        feRefComercial.TextMatrix(I + 1, 3) = oPersona.ObtenerRefComFono(I) 'Telefono de la Referencia
        feRefComercial.TextMatrix(I + 1, 4) = oPersona.ObtenerRefComNumRef(I) 'Número de Referencia
    Next I
    
    feRefComercial.lbEditarFlex = False
    
End Sub

Private Sub CargaRefBancarias()
Dim I As Integer

    feRefBancaria.lbEditarFlex = True
    Call LimpiaFlex(feRefBancaria)
    For I = 0 To oPersona.NumeroRefBancaria - 1
        feRefBancaria.AdicionaFila
        feRefBancaria.TextMatrix(I + 1, 1) = oPersona.ObtenerRefBanCodIF(I) 'Codigo Institución Financiera
        feRefBancaria.TextMatrix(I + 1, 2) = oPersona.ObtenerRefBanNombre(I) 'Tipo de Relacion con la Referencia
        feRefBancaria.TextMatrix(I + 1, 3) = oPersona.ObtenerRefBanNumCta(I) 'Número de Cuenta
        feRefBancaria.TextMatrix(I + 1, 4) = oPersona.ObtenerRefBanNumTar(I) 'Número de Tarjeta
        feRefBancaria.TextMatrix(I + 1, 5) = Format$(oPersona.ObtenerRefBanLinCred(I), "###,###.00") 'Monto de la Línea de Crédito
    Next I
    
    feRefBancaria.lbEditarFlex = False
    
End Sub

Private Sub CargaPatVehicular()
Dim I As Integer

    fePatVehicular.lbEditarFlex = True
    Call LimpiaFlex(fePatVehicular)
    For I = 0 To oPersona.NumeroPatVehicular - 1
        fePatVehicular.AdicionaFila
        fePatVehicular.TextMatrix(I + 1, 1) = oPersona.ObtenerPatVehMarca(I) 'Marca del Patrimonio Vehicular
        fePatVehicular.TextMatrix(I + 1, 2) = oPersona.ObtenerPatVehFecFab(I) 'Fecha de Fabricacion
        fePatVehicular.TextMatrix(I + 1, 3) = Format$(oPersona.ObtenerPatVehValCom(I), "###,###.00") 'Valor Comercial
        fePatVehicular.TextMatrix(I + 1, 4) = oPersona.ObtenerPatVehCondicion(I) 'Condicion del Patrimonio
        fePatVehicular.TextMatrix(I + 1, 5) = oPersona.ObtenerPatVehCod(I) 'Codigo del Patrimonio Vehicular
    Next I
    
    fePatVehicular.lbEditarFlex = False
    
End Sub

Private Sub CargaFuentesIngreso()
Dim I As Integer
Dim MatFte As Variant

    Call LimpiaFlex(FEFteIng)
    
    MatFte = oPersona.FiltraFuentesIngresoPorRazonSocial
    
    'For i = 0 To oPersona.NumeroFtesIngreso - 1
    '    FEFteIng.AdicionaFila
    '    FEFteIng.TextMatrix(i + 1, 0) = i + 1
    '    FEFteIng.TextMatrix(i + 1, 1) = IIf(oPersona.ObtenerFteIngTipo(i) = "1", "D", "I") 'Tipo de Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 2) = oPersona.ObtenerFteIngRazonSocial(i) 'Razon Social de Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 3) = Format(oPersona.ObtenerFteIngFecCaducac(i), "dd/mm/yyyy") 'Fecha de Caducacion de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 4) = Format(oPersona.ObtenerFteIngFecEval(i), "dd/mm/yyyy") 'Fecha de Evaluacion de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 5) = IIf(oPersona.ObtenerFteIngMoneda(i) = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 5) = IIf(oPersona.ObtenerFteIngMoneda(i) = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
    'Next i
    
    If IsArray(MatFte) Then
        For I = 0 To UBound(MatFte) - 1
            FEFteIng.AdicionaFila
            FEFteIng.TextMatrix(I + 1, 0) = MatFte(I, 0)
            FEFteIng.TextMatrix(I + 1, 1) = MatFte(I, 1)
            FEFteIng.TextMatrix(I + 1, 2) = MatFte(I, 2)
            FEFteIng.TextMatrix(I + 1, 3) = MatFte(I, 5)
            FEFteIng.TextMatrix(I + 1, 4) = MatFte(I, 6)
            FEFteIng.TextMatrix(I + 1, 5) = MatFte(I, 7)
        Next I
    End If
End Sub
Private Sub CargaDatos()
Dim I As Integer
     
    bEstadoCargando = True
    SSTDatosGen.Tab = 0
    SSTabs.Tab = 0
    'Carga Personeria
    cmbPersPersoneria.ListIndex = IndiceListaCombo(cmbPersPersoneria, Trim(Str(oPersona.Personeria)))
    
    'Habilita o Deshabilita Ficha de Persona Juridica
    
    Call HabilitaFichaPersonaJur(False)
    Call HabilitaFichaPersonaNat(False)
    
    'Carga Ubicacion Georgrafica
    If Len(Trim(oPersona.UbicacionGeografica)) = 12 Then
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "04028")
        cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbicacionGeografica, 2, 2) & String(9, "0"))
        cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbicacionGeografica, 2, 4) & String(7, "0"))
        cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbicacionGeografica, 2, 6) & String(5, "0"))
        cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & oPersona.UbicacionGeografica)
    Else
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & oPersona.UbicacionGeografica)
        cmbPersUbiGeo(1).Clear
        cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(1).ListIndex = 0
        cmbPersUbiGeo(2).Clear
        cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(2).ListIndex = 0
        cmbPersUbiGeo(3).Clear
        cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(3).ListIndex = 0
        cmbPersUbiGeo(4).Clear
        cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(4).ListIndex = 0
    End If
    
    'Carga Direccion
    txtPersDireccDomicilio.Text = oPersona.Domicilio
    
    'Selecciona la Condicion del Doicilio
    cmbPersDireccCondicion.ListIndex = IndiceListaCombo(cmbPersDireccCondicion, oPersona.CondicionDomicilio)
    
    txtValComercial.Text = Format(oPersona.ValComDomicilio, "#####,###.00")
    
    'Tipo de sangre
    If CboTipoSangre.ListCount > 0 Then
        CboTipoSangre.ListIndex = IndiceListaCombo(CboTipoSangre, oPersona.TipoSangre)
    End If
    
    'Carga Ficha 1
    If oPersona.Personeria = gPersonaNat Then
        txtPersNombreAP.Text = oPersona.ApellidoPaterno
        txtPersNombreAM.Text = oPersona.ApellidoMaterno
        txtPersNombreN.Text = oPersona.nombres
    Else
        txtPersNombreRS.Text = oPersona.NombreCompleto
    End If
    TxtTalla.Text = Format(oPersona.Talla, "#0.00")
    TxtPeso.Text = Format(oPersona.Peso, "#0.00")
    TxtEmail.Text = oPersona.Email
    
    txtPersTelefono2.Text = oPersona.Telefonos2
    'EJVG20111212 ***********************
    txtCel1.Text = oPersona.Celular
    txtCel2.Text = oPersona.Celular2
    txtCel3.Text = oPersona.Celular3
    TxtEmail2.Text = oPersona.Email2
    
    If oPersona.Sexo = "F" Then
        TxtApellidoCasada.Text = oPersona.ApellidoCasada
        cmbPersNatSexo.ListIndex = 0
        Call DistribuyeApellidos(True)
    Else
        cmbPersNatSexo.ListIndex = 1
        Call DistribuyeApellidos(False)
    End If
    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, oPersona.EstadoCivil)
    txtPersNatHijos.Text = Trim(Str(oPersona.Hijos))
    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, Space(30) & oPersona.Nacionalidad)
    chkResidente.value = oPersona.Residencia
    
    'Carga Datos Generales
    txtPersNacCreac.Text = Format(oPersona.FechaNacimiento, "dd/mm/yyyy")
    txtPersTelefono.Text = oPersona.Telefonos
    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, oPersona.CIIU)
    
    Call CargaControlEstadoPersona(oPersona.Personeria)
    cmbPersEstado.ListIndex = IndiceListaCombo(cmbPersEstado, Right("0" & oPersona.Estado, 2))
    
    TxtSiglas.Text = oPersona.Siglas     'Carga Razon Social
    
    TxtSbs.Text = oPersona.PersCodSbs    'Carga Codigo SBS
    
    'Selecciona el Tipo de Persona Juridica
    cmbPersJurTpo.ListIndex = IndiceListaCombo(cmbPersJurTpo, Trim(Str(IIf(oPersona.TipoPersonaJur = "", -1, oPersona.TipoPersonaJur))))

    'Selecciona la relacion Con la Persona
    CmbRela.ListIndex = IndiceListaCombo(CmbRela, oPersona.PersRelInst)

    'Selecciona la magnitud Empresarial
    cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, Trim(oPersona.MagnitudEmpresarial))
    
    'Carga Numero de Empleados
    txtPersJurEmpleados.Text = Trim(Str(oPersona.NumerosEmpleados))
    
    Call CargaDocumentos                    'Carga Los Documentos de la Personas
    
    Call CargaRelacionesPersonas            'Carga las Relaciones de las Personas
    
    Call CargaFuentesIngreso                'Carga las Fuentes de Ingresos de las Personas
    
    Call CargaRefComerciales                'Carga las Referencias Comerciales
    lnNumRefCom = oPersona.MaxRefComercial  'Carga el max Ref Comercial
     
    Call CargaRefBancarias                  'Carga las Referencias Bancarias
    
    Call CargaPatVehicular                  'Carga el Patrimonio Vehicular
    lnNumPatVeh = oPersona.MaxPatVehicular  'Carga el max Pat Vehicular
    
    'Carga Firma
    Call IDBFirma.CargarFirma(oPersona.RFirma)
    
    If Not oPersona Is Nothing Then
    If oPersona.Personeria = gPersonaNat Then
        SSTabs.Tab = 0
    Else
        SSTabs.Tab = 1
    End If
    End If
    bEstadoCargando = False
    CmdPersFteConsultar.Enabled = True
     'WIOR 20130827 *******************************
    fsNombreActual = oPersona.NombreCompleto
    Set rsDocPersActual = FEDocs.GetRsNew
    'WIOR FIN *************************************
    
    'ARLO 20160126 ***
    gsopecod = LogPistaConsultaPersona
    Set objPista = New COMManejador.Pista
    'LUCV20181220, 'LUCV20181220 Agregó gConsultar, Anexo01 de Acta 199-2018
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Consulta del cliente : " & txtPersNombreAP.Text & " " & txtPersNombreAM & " " & txtPersNombreN.Text, TxtBCodPers.Text, 3
    Set objPista = Nothing
    '***
End Sub

Private Sub CargaUbicacionesGeograficas()
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim I As Integer

On Error GoTo ErrCargaUbicacionesGeograficas
    Set Conn = New DConecta
    'Carga Niveles
    sSql = "select *, 1 p from UbicacionGeografica where cUbiGeoCod like '0%' "
    sSql = sSql & " Union "
    sSql = sSql & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
    sSql = sSql & " Union "
    sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    ContNiv4 = 0
    ContNiv5 = 0
    
    Conn.AbreConexion
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        Select Case R!P
            Case 1 'Pais
                ContNiv1 = ContNiv1 + 1
                ReDim Preserve Nivel1(ContNiv1)
                Nivel1(ContNiv1 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 2 ' Departamento
                ContNiv2 = ContNiv2 + 1
                ReDim Preserve Nivel2(ContNiv2)
                Nivel2(ContNiv2 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 3 'Provincia
                ContNiv3 = ContNiv3 + 1
                ReDim Preserve Nivel3(ContNiv3)
                Nivel3(ContNiv3 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 4 'Distrito
                ContNiv4 = ContNiv4 + 1
                ReDim Preserve Nivel4(ContNiv4)
                Nivel4(ContNiv4 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 5 'Zona
                ContNiv5 = ContNiv5 + 1
                ReDim Preserve Nivel5(ContNiv5)
                Nivel5(ContNiv5 - 1) = Trim(R!cUbigeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
        End Select
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For I = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(I)
        cmbNacionalidad.AddItem Nivel1(I)
        If Trim(Right(Nivel1(I), 10)) = "04028" Then
            nPos = I
        End If
    Next I
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    cmbPersUbiGeo(4).Clear
    cmbNacionalidad.ListIndex = nPos
    Exit Sub
    
ErrCargaUbicacionesGeograficas:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TTipoCombo)
Dim I As Long
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboDpto
            cmbPersUbiGeo(1).Clear
            If sCodigo = "04028" Then
                For I = 0 To ContNiv2 - 1
                    cmbPersUbiGeo(1).AddItem Nivel2(I)
                Next I
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboProv
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(2).Clear
                For I = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel3(I), 15)), 2, 2) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(I)
                    End If
                Next I
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For I = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel4(I), 15)), 2, 4) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(I)
                    End If
                Next I
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(4).Clear
            If Len(sCodigo) > 3 Then
                For I = 0 To ContNiv5 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel5(I), 15)), 2, 6) Then
                        cmbPersUbiGeo(4).AddItem Nivel5(I)
                    End If
                Next I
            Else
                cmbPersUbiGeo(4).AddItem psValor
            End If
    End Select
End Sub

Private Sub CboPersCiiu_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 20))
    End If
End Sub

Private Sub CboPersCiiu_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 10))
    End If
End Sub

Private Sub CboPersCiiu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPersEstado.SetFocus
    End If
End Sub

Private Sub CboTipoSangre_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoSangre = Trim(Right(CboTipoSangre.Text, 2))
    End If
End Sub

Private Sub CboTipoSangre_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoSangre = Trim(Right(CboTipoSangre.Text, 2))
    End If
End Sub

Private Sub CboTipoSangre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPeso.SetFocus
    End If
End Sub

Private Sub chkResidente_Click()

    On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Residencia = chkResidente.value
    End If
    If TxtTalla.Enabled Then TxtTalla.SetFocus
End Sub

Private Sub chkResidente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    On Error Resume Next
        If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Residencia = chkResidente.value
    End If
    If TxtTalla.Enabled Then TxtTalla.SetFocus
End If
End Sub

Private Sub cmbNacionalidad_Change()

    On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
    End If
    
End Sub

Private Sub cmbNacionalidad_Click()
    On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
        If Trim(Right(cmbNacionalidad.Text, 12)) = "04028" Then
            chkResidente.value = 1
        End If
    End If
End Sub

Private Sub cmbNacionalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkResidente.SetFocus
End If
End Sub

Private Sub cmbPersDireccCondicion_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersDireccCondicion_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersDireccCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValComercial.SetFocus
    End If
End Sub

Private Sub cmbPersEstado_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))
      End If
End Sub

Private Sub cmbPersEstado_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))
    End If
End Sub

Private Sub cmbPersEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbRela.SetFocus
    End If
End Sub

Private Sub cmbPersJurMagnitud_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
            oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
        End If
    End If
End Sub

Private Sub cmbPersJurMagnitud_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
    End If
End Sub

Private Sub cmbPersJurMagnitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNacCreac.SetFocus
    End If
End Sub

Private Sub cmbPersJurTpo_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
    End If
End Sub

Private Sub cmbPersJurTpo_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
    End If
End Sub

Private Sub cmbPersJurTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersJurEmpleados.SetFocus
    End If
End Sub

Private Sub cmbPersNatEstCiv_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
    
  End If
End Sub

Private Sub cmbPersNatEstCiv_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            TxtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersNatEstCiv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNatHijos.SetFocus
    End If
End Sub

Private Sub cmbPersNatSexo_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
  End If
End Sub

Private Sub cmbPersNatSexo_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
        If oPersona.EstadoCivil = "" Then
            Exit Sub
        End If
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            TxtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersNatSexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPersNatEstCiv.SetFocus
    End If
End Sub

Private Sub cmbPersPersoneria_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersPersoneria.Text, 15))
    End If
End Sub

Private Sub cmbPersPersoneria_Click()
    If Not bEstadoCargando Then
        If oPersona.Personeria <> CInt(Trim(Right(cmbPersPersoneria.Text, 15))) Then
            oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", Trim(Str(gPersonaNat)), cmbPersPersoneria.Text), 15))
            If oPersona.Personeria <> gPersonaNat Then
                Call HabilitaFichaPersonaJur(True)
                Call HabilitaFichaPersonaNat(False)
                SSTabs.Tab = 1
            Else
                Call HabilitaFichaPersonaJur(False)
                Call HabilitaFichaPersonaNat(True)
                SSTabs.Tab = 0
            End If
            cmbPersEstado.ListIndex = -1
        End If
    End If
    Call CargaControlEstadoPersona(oPersona.Personeria)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", Trim(Str(gPersonaNat)), cmbPersPersoneria.Text), 15))
    End If
End Sub

Private Sub cmbPersPersoneria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPersNombreAP.Enabled Then
            txtPersNombreAP.SetFocus
        Else
            txtPersNombreRS.SetFocus
        End If
    End If
End Sub

Private Sub cmbPersUbiGeo_Change(Index As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
        Select Case Index
            Case 0 'Combo Pais
                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboDpto)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(2).Clear
                    cmbPersUbiGeo(3).Clear
                    cmbPersUbiGeo(4).Clear
                End If
            Case 1 'Combo Dpto
                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboProv)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(3).Clear
                    cmbPersUbiGeo(4).Clear
                End If
            Case 2 'Combo Provincia
                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboDist)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(4).Clear
                End If
            Case 3 'Combo Distrito
                Call ActualizaCombo(cmbPersUbiGeo(3).Text, ComboZona)
        End Select
        
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
        End If

End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            txtPersDireccDomicilio.SetFocus
        End If
    End If
End Sub

Private Sub CmbRela_Change()
    Call CmbRela_Click
End Sub

Private Sub CmbRela_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PersRelInst = CInt(Trim(Right(CmbRela.Text, 10)))
    End If
End Sub

Private Sub CmbRela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTDatosGen.Tab = 1
        cmbPersUbiGeo(0).SetFocus
    End If
End Sub

Private Sub CmdActFirma_Click()
Dim sRuta As String
    CdlgImg.nHwd = Me.hwnd
    CdlgImg.Show
    sRuta = CdlgImg.Ruta
    If Len(Trim(sRuta)) > 0 Then
        IDBFirma.RutaImagen = sRuta
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
        Else
            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
        End If
    End If
End Sub

Private Sub cmdEditar_Click()
    If oPersona Is Nothing Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.PersCodigo = "" Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    CmdPersAceptar.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Enabled = True
    CmdPersCancelar.Visible = True
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    
    Call HabilitaControlesPersona(True)
    bNuevaPersona = False
    bPersonaAct = True
    
    '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
    If gnPersPersoneria = 1 And gsCodArea <> "022" Then

        FERelPers.lbEditarFlex = False
        cmdPersRelacNew.Enabled = False
        cmdPersRelacEditar.Enabled = False
        cmdPersRelacDel.Enabled = False
        cmdPersRelacAceptar.Enabled = False
        cmdPersRelacCancelar.Enabled = False
    
    End If
    
    '***Fin Modificado por ELRO
    'EJVG20111217 **********************************
    bPermisoEditarTodo = True
    If oPersona.Personeria = gPersonaNat Then
        If ObtenerVecesCreditoyAhorroPersona(oPersona.PersCodigo) > 0 Then
            If Not validaPermisoEditarPersona(gsCodCargo) Then
                bPermisoEditarTodo = False
                HabilitarControlesDatosPrincipalesPersonas (bPermisoEditarTodo)
            End If
        End If
    End If
    'WIOR 20130827 *********************************************
     Dim oGen As UAcceso
     Set oGen = New UAcceso
     fbPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, 3)
     If Not bPermisoEditarTodo Or oPersona.Personeria <> gPersonaNat Then
          HabilitarControlesDatosBasicos (fbPermisoCargo)
     End If
     If bPermisoEditarTodo Then
        fbPermisoCargo = True
     End If
    'WIOR FIN **************************************************
End Sub

Private Sub HabilitarControlesDatosPrincipalesPersonas(ByVal pbHabilita As Boolean)
    cmbPersNatSexo.Enabled = pbHabilita
    cmbPersNatEstCiv.Enabled = pbHabilita
    txtPersNombreAP.Enabled = pbHabilita
    txtPersNombreAM.Enabled = pbHabilita
    TxtApellidoCasada.Enabled = pbHabilita
    txtPersNombreN.Enabled = pbHabilita
    txtPersNacCreac.Enabled = pbHabilita
    cmdPersIDAceptar.Enabled = pbHabilita
    cmdPersIDedit.Enabled = pbHabilita
    cmdPersIDDel.Enabled = pbHabilita
End Sub
'WIOR 20130826 *****TI-ERS119-2013**********************************************
Private Sub HabilitarControlesDatosBasicos(ByVal pbHabilita As Boolean)
If oPersona.Personeria = gPersonaNat Then
    txtPersNombreAP.Enabled = pbHabilita
    txtPersNombreAM.Enabled = pbHabilita
    TxtApellidoCasada.Enabled = pbHabilita
    txtPersNombreN.Enabled = pbHabilita
Else
    txtPersNombreRS.Enabled = pbHabilita
End If
cmdPersIDAceptar.Enabled = pbHabilita
cmdPersIDedit.Enabled = pbHabilita
cmdPersIDDel.Enabled = pbHabilita
End Sub
'WIOR FIN *******************************************************************
Private Sub CmdFteIngEditar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Registros para Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.NumeroFtesIngreso > 0 Then
        Call frmFteIngresos.Editar(FEFteIng.row - 1, oPersona)
    Else
        MsgBox "No Existe Fuentes de Ingreso para Editar", vbInformation, "Aviso"
    End If
End Sub


Private Sub CmdFteIngEliminar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Registros para Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta Seguro que Desea Eliminar esta Fuente de Ingreso", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarFteIngTipoAct(PersFilaEliminda, FEFteIng.row - 1)
        FEFteIng.EliminaFila FEFteIng.row
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            Call CmdPersAceptar_Click
        Else
            oPersona.LimpiaEliminados
        End If
        Call CargaFuentesIngreso
    End If
End Sub

Private Sub CmdFteIngNuevo_Click()
Dim sNombreCompleto As String
    
    'If bNuevaPersona Then
    '    MsgBox "Se debe Crear la Persona "
    'End If
    
    If Not ValidaControles Then
        Exit Sub
    End If
    If Trim(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
        If Trim(Right(cmbPersNatSexo.Text, 2)) <> "F" And Len(Trim(TxtApellidoCasada.Text)) > 0 Then
            sNombreCompleto = txtPersNombreAP.Text & "/" & txtPersNombreAM.Text & "\" & TxtApellidoCasada.Text & "," & txtPersNombreN.Text
        Else
            sNombreCompleto = txtPersNombreAP.Text & "/" & txtPersNombreAM.Text & "," & txtPersNombreN.Text
        End If
    Else
        sNombreCompleto = txtPersNombreRS.Text
    End If
    oPersona.NombreCompleto = sNombreCompleto
    If Trim(FEFteIng.TextMatrix(1, 1)) = "" Then
        Call frmFteIngresos.NuevaFteIngreso(oPersona)
    Else
        'Call frmFteIngresos.NuevaFteIngreso(oPersona, CInt(Trim(FEFteIng.TextMatrix(FEFteIng.Row, 5))))
    End If
    Call CargaFuentesIngreso
End Sub

Private Sub CmdNuevo_Click()
    
    Call HabilitaControlesPersona(True)
    Call LimpiarPantalla
    bEstadoCargando = True
    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, Trim(Str(gPersEstadoCivilCasado)))
    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, "04028")
    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, "O9309")
    chkResidente.value = 1
    bEstadoCargando = False
    'Inicializa Controles
    TxtBCodPers.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTabs.Enabled = True
    SSTabs.Tab = 0
    SSTDatosGen.Enabled = True
    SSTDatosGen.Tab = 0
    SSTIdent.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    If oPersona Is Nothing Then
       Set oPersona = New DPersona
    End If
    Call oPersona.NuevaPersona
    oPersona.TipoActualizacion = PersFilaNueva
    
    Call HabilitaControlesPersonaFtesIngreso(True)
    cmbPersNatSexo.ListIndex = 0
    cmbPersPersoneria.ListIndex = 0
        
    If Not bBuscaNuevo Then
        cmbPersPersoneria.SetFocus
    End If
    If txtPersNombreAP.Enabled And txtPersNombreAP.Visible Then
        txtPersNombreAP.SetFocus
    End If
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
    oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 10))
    oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
    oPersona.Residencia = chkResidente.value
    bNuevaPersona = True
    bPersonaAct = True
    
    FEFteIng.Enabled = False
    CmdFteIngNuevo.Enabled = False
    CmdFteIngEditar.Enabled = False
    CmdFteIngEliminar.Enabled = False
    CmdPersFteConsultar.Enabled = False
    TxtBCodPers.Text = ""
    'WIOR 20130827 ***************************
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
End Sub

Private Sub cmdPatVehAcepta_Click()

    If Not ValidaDatosPatVehicular Then
        Exit Sub
    End If
    
   If cmdPersPatVehicularEjecutado = 1 Then
        Call oPersona.AdicionaPatVehicular
        Call oPersona.ActualizarPatVehTipoAct(PersFilaNueva, fePatVehicular.row - 1)
        lnNumPatVeh = lnNumPatVeh + 1
        fePatVehicular.TextMatrix(fePatVehicular.row, 5) = lnNumPatVeh
    Else
        If cmdPersPatVehicularEjecutado = 2 Then
            If oPersona.ObtenerPatVehTipoAct(fePatVehicular.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarPatVehTipoAct(PersFilaModificada, fePatVehicular.row - 1)
            End If
        End If
    End If
    
    'Marca
    Call oPersona.ActualizaPatVehMarca(UCase(fePatVehicular.TextMatrix(fePatVehicular.row, 1)), fePatVehicular.row - 1)
    'Fecha Fabricación
    If fePatVehicular.TextMatrix(fePatVehicular.row, 2) <> "" Then
        Call oPersona.ActualizaPatVehFecFab(CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)), fePatVehicular.row - 1)
    End If
    'Valor de Comercializacion
    Call oPersona.ActualizaPatVehValCom(fePatVehicular.TextMatrix(fePatVehicular.row, 3), fePatVehicular.row - 1)
    'Condicion del Vehiculo
    Call oPersona.ActualizaPatVehCondicion(fePatVehicular.TextMatrix(fePatVehicular.row, 4), fePatVehicular.row - 1)
    Call oPersona.ActualizaPatVehCod(fePatVehicular.TextMatrix(fePatVehicular.row, 5), fePatVehicular.row - 1)
            
    'Habilitar Controles
    cmdPersPatVehicularEjecutado = 0
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    fePatVehicular.SetFocus

End Sub

Private Sub cmdPatVehCancela_Click()
    CargaPatVehicular
    'Habilitar Controles
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    
    fePatVehicular.SetFocus

End Sub

Private Sub cmdPatVehEdita_Click()
   
   If oPersona.NumeroPatVehicular > 0 Then
        cmdPersPatVehicularEjecutado = 2
        FEPatVehPersNoMoverdeFila = fePatVehicular.row
        NomMoverSSTabs = SSTabs.Tab
        fePatVehicular.lbEditarFlex = True
        fePatVehicular.SetFocus
        cmdPatVehNuevo.Enabled = False
        cmdPatVehEdita.Enabled = False
        cmdPatVehElimina.Enabled = False
        cmdPatVehAcepta.Visible = True
        cmdPatVehCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        fePatVehicular.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdPatVehElimina_Click()
    If MsgBox("Esta Seguro que Desea Eliminar el Patrimonio Vehicular", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarPatVehTipoAct(PersFilaEliminda, fePatVehicular.row - 1)
        Call CmdPersAceptar_Click
        Call CargaPatVehicular
    End If
End Sub

Private Sub cmdPatVehNuevo_Click()
    cmdPatVehAcepta.Visible = True
    cmdPatVehCancela.Visible = True
    cmdPatVehNuevo.Enabled = False
    cmdPatVehElimina.Enabled = False
    cmdPatVehEdita.Enabled = False
    fePatVehicular.lbEditarFlex = True
    fePatVehicular.AdicionaFila
    cmdPersPatVehicularEjecutado = 1
    FEPatVehPersNoMoverdeFila = fePatVehicular.Rows - 1
    fePatVehicular.SetFocus
End Sub

Private Sub CmdPersAceptar_Click()
Dim oPersonaNeg As nPersona
Dim R As ADODB.Recordset

    If Not ValidaControles Then
        Exit Sub
    End If
    
    'Verifica Homonimia
    If oPersona.TipoActualizacion = PersFilaNueva Then
        Set oPersonaNeg = New nPersona
        Set R = oPersonaNeg.VerificandoHomonimia(oPersona)
        If Not R.BOF And Not R.EOF Then
            Call frmMuestraHomonimia.Inicio(R)
        End If
        R.Close
        Set R = Nothing
        Set oPersonaNeg = Nothing
    End If

    If MsgBox("Se va a proceder a guardar los datos de la Persona, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
        
    
    On Error GoTo ErrorCmdPersAceptar
    
    'Verificando Duplicidad de Documento
    Set oPersonaNeg = New nPersona
    If oPersonaNeg.VerificandoDuplicidadDoc(oPersona) <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.row, 1), 30)) & " se Encuentra Duplicado", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Set oPersonaNeg = Nothing
        Exit Sub
    End If
    Set oPersonaNeg = Nothing
    
    'Verificando Tamaño de Documento
    Set oPersonaNeg = New nPersona
    If oPersonaNeg.VerificandoTamanioDoc(oPersona) <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.row, 1), 30)) & " numero de digitos Incorrecto", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Set oPersonaNeg = Nothing
        Exit Sub
    End If
    Set oPersonaNeg = Nothing
    Screen.MousePointer = 11
    
    'Genera el Nuevo codigo de la Persona
    'If oPersona.TipoActualizacion = PersFilaNueva Then
        'Set oPersonaNeg = New NPersona
        'oPersona.PersCodigo = oPersonaNeg.GenerarNuevoCodigoPersona(oPersona, gsCodCMAC & gsCodAge)
        'Set oPersonaNeg = Nothing
    'End If
    
    oPersona.CampoActualizacion = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    oPersona.dFechaHoy = gdFecSis
    Call oPersona.GrabarCambiosPersona(gsCodCMAC, gsCodAge)
    Call HabilitaControlesPersona(False)
    TxtBCodPers.Text = oPersona.PersCodigo
    'WIOR 20130827 **************************************
    If fbPermisoCargo Then
        Dim cMovCambio As String
        Dim oCambio As DPersonas
        Dim bCabecera As Boolean
        Dim nI, nJ, nZ, nAux As Integer
        Set oCambio = New DPersonas
        Dim oNContFunciones As NContFunciones
        Set oNContFunciones = New NContFunciones
        Dim nTpoDocAnt, nTpoDocAct As Integer
        Dim cNumDocAnt, cNumDocAct As String
        bCabecera = False
        Set rsDocPersUlt = FEDocs.GetRsNew
        nI = rsDocPersActual.RecordCount
        nJ = rsDocPersUlt.RecordCount
        
        nZ = IIf(nI > nJ, nI, nJ)
        
        cMovCambio = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        
        If Trim(fsNombreActual) <> Trim(oPersona.NombreCompleto) Then
            Call oCambio.RegistroDatosCambiosDatosPrinc(False, cMovCambio, oPersona.PersCodigo, Trim(fsNombreActual), Trim(oPersona.NombreCompleto))
            bCabecera = True
        End If
        
        For nAux = 1 To nZ
            If nI >= nAux Then
                nTpoDocAnt = CInt(Trim(Right(rsDocPersActual!Tipo, 5)))
                cNumDocAnt = Trim(rsDocPersActual!Numero)
                rsDocPersActual.MoveNext
            Else
                nTpoDocAnt = 0
                cNumDocAnt = ""
            End If
            
            If nJ >= nAux Then
                nTpoDocAct = CInt(Trim(Right(rsDocPersUlt!Tipo, 5)))
                cNumDocAct = Trim(rsDocPersUlt!Numero)
                rsDocPersUlt.MoveNext
            Else
                nTpoDocAct = 0
                cNumDocAct = ""
            End If
            
            If (nTpoDocAnt <> nTpoDocAct) Or (cNumDocAnt <> cNumDocAct) Then
                Call oCambio.RegistroDatosCambiosDatosPrinc(True, cMovCambio, oPersona.PersCodigo, , , nTpoDocAnt, nTpoDocAct, cNumDocAnt, cNumDocAct)
                bCabecera = True
            End If
        Next nAux
        
        If bCabecera Then
            Call oCambio.RegistroCabeceraDatosCambiosDatosPrinc(cMovCambio, oPersona.PersCodigo, 2)
        End If
    End If
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN *******************************************
    Screen.MousePointer = 0
    
    MsgBox "Datos Grabados", vbInformation, "Aviso"
    Call HabilitaControlesPersona(False)
    CmdPersAceptar.Visible = False
    CmdPersCancelar.value = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    NomMoverSSTabs = -1
    
    'Habilita Todos los Controles
    SSTabs.Enabled = True
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    cmdNuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
    
    bPersonaAct = False
    
    If bBuscaNuevo Then
        sPersCodNombre = oPersona.PersCodigo & oPersona.NombreCompleto
        Unload Me
    End If
    Call TxtBCodPers_EmiteDatos  'WIOR 20130828
    Exit Sub
    
ErrorCmdPersAceptar:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Call CmdPersCancelar_Click
End Sub

Private Sub CmdPersCancelar_Click()
    'Habilita Todos los Controles
    NomMoverSSTabs = -1
    SSTabs.Enabled = True
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    
    If Trim(TxtBCodPers.Text) <> "" Then
        Call TxtBCodPers_EmiteDatos
    Else
        TxtBCodPers.Text = ""
        Call LimpiarPantalla
        Call HabilitaControlesPersona(False)
        Call HabilitaControlesPersonaFtesIngreso(False)
        Call HabilitaFichaPersonaJur(False)
        Call HabilitaFichaPersonaNat(False)
    End If
    cmdNuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
    
    If Me.cmdNuevo.Visible And Me.cmdNuevo.Enabled Then
        cmdNuevo.SetFocus
    ElseIf Me.cmdSalir.Visible And Me.cmdSalir.Enabled Then
        cmdSalir.SetFocus
    Else
        If Me.Visible Then Me.SetFocus
    End If
    bPersonaAct = False
    'EJVG20111219
    Me.cmdPersIDAceptar.Visible = False
    Me.cmdPersIDCancelar.Visible = False
    Me.cmdPersIDnew.Visible = True
    Me.cmdPersIDedit.Visible = True
    If bBuscaNuevo Then
        sPersCodNombre = ""
        Unload Me
    End If
    'WIOR 20130827 ***************************
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
End Sub

Private Sub CmdPersFteConsultar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Fuentes de Ingreso para Consultar", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.NumeroFtesIngreso > 0 Then
        Call frmFteIngresos.ConsultarFuenteIngreso(FEFteIng.row - 1, oPersona)
    Else
        MsgBox "No se puede Consultar la Fuente de Ingreso", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdPersIDAceptar_Click()

    If Not ValidaDatosDocumentos Then
        Exit Sub
    End If
    
    If cmdPersDocEjecutado = 1 Then
        Call oPersona.AdicionaDocumento(PersFilaNueva, FEDocs.TextMatrix(FEDocs.row, 2), FEDocs.TextMatrix(FEDocs.row, 1))
        Call oPersona.ActualizarDocsTipoAct(PersFilaNueva, FEDocs.row - 1)
    Else
        If cmdPersDocEjecutado = 2 Then
            If oPersona.ObtenerDocTipoAct(FEDocs.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarDocsTipoAct(PersFilaModificada, FEDocs.row - 1)
            End If
        End If
    End If
    
    'Tipo de Docmumento
    Call oPersona.ActualizaDocsTipo(FEDocs.TextMatrix(FEDocs.row, 1), FEDocs.row - 1)
    'Tipo de Numero
    Call oPersona.ActualizaDocsNumero(FEDocs.TextMatrix(FEDocs.row, 2), FEDocs.row - 1)
    
    'Habilitar Controles
    cmdPersDocEjecutado = 0
    FEDocsPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    cmdPersIDnew.Visible = True
    'cmdPersIDedit.Enabled = True
    If bPermisoEditarTodo = True Then 'EJVG20111219
        cmdPersIDedit.Enabled = True
        cmdPersIDDel.Enabled = True
    End If
    'WIOR 20130827 ******************************
    If Not bPermisoEditarTodo Then
        If fbPermisoCargo Then
            cmdPersIDedit.Enabled = True
            cmdPersIDDel.Enabled = True
        End If
    End If
    'WIOR FIN ***********************************
    cmdPersIDedit.Visible = True
    'cmdPersIDDel.Enabled = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    'cmdNuevo.Enabled = True 'WIOR 20130827 COMENTÓ
    'cmdEditar.Enabled = True 'WIOR 20130827 COMENTÓ
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Enabled = True
        CmdPersCancelar.Visible = True
    End If
    
    If oPersona.TipoActualizacion = PersFilaNueva Then
        'SSTabs.Tab = 3
        'If CmdFteIngNuevo.Enabled Then
        '    CmdFteIngNuevo.SetFocus
        'End If
        CmdPersAceptar.SetFocus
        
    Else
        FEDocs.SetFocus
    End If
    
End Sub

Private Sub cmdPersIDCancelar_Click()

    Call CargaDocumentos
    
    'Habilitar Controles
    FEDocsPersNoMoverdeFila = -1
    cmdPersDocEjecutado = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    'cmdPersIDedit.Enabled = True
    'cmdPersIDDel.Enabled = True
    If bPermisoEditarTodo = True Then 'EJVG20111219
        cmdPersIDedit.Enabled = True
        cmdPersIDDel.Enabled = True
    End If
    'WIOR 20130827 ******************************
    If Not bPermisoEditarTodo Then
        If fbPermisoCargo Then
            cmdPersIDedit.Enabled = True
            cmdPersIDDel.Enabled = True
        End If
    End If
    'WIOR FIN ***********************************
    cmdPersIDnew.Visible = True
    cmdPersIDedit.Visible = True
    cmdPersIDDel.Visible = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    'cmdNuevo.Enabled = True 'WIOR 20130827 COMENTÓ
    'cmdEditar.Enabled = True 'WIOR 20130827 COMENTÓ
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    If FEDocs.Rows <= 2 Then    'ACA MODIFIQUE
        cmdPersIDDel.Enabled = False
    End If
    
    FEDocs.SetFocus

End Sub

Private Sub cmdPersIDDel_Click()
    If FEDocs.Rows <= 2 And Trim(FEDocs.TextMatrix(1, 0)) = "" Then
        MsgBox " No existe ningun documento para eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta Seguro que Desea Eliminar este Documento", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            Call oPersona.ActualizarDocsTipoAct(PersFilaEliminda, FEDocs.row - 1)
            Call CmdPersAceptar_Click
            Call CargaDocumentos
        Else
            Call oPersona.EliminarDocumento(CInt(Trim(Right(FEDocs.TextMatrix(FEDocs.row, 1), 2))), Trim(FEDocs.TextMatrix(FEDocs.row, 2)))
            If FEDocs.Rows > 1 Then
                Call FEDocs.EliminaFila(FEDocs.row)
            End If
            
        End If
        
        If FEDocs.Rows <= 2 Then 'ACA MODIFIQUE
            cmdPersIDDel.Enabled = False
        End If
    
    End If
End Sub

Private Sub cmdPersIDedit_Click()
    cmdPersDocEjecutado = 2
    FEDocsPersNoMoverdeFila = FEDocs.row
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
        
End Sub

Private Sub cmdPersIDnew_Click()
    FEDocs.AdicionaFila
    cmdPersDocEjecutado = 1
    FEDocsPersNoMoverdeFila = FEDocs.Rows - 1
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
End Sub

Private Sub cmdPersRelacAceptar_Click()
    
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 5))) = 0 Then
        FERelPers.TextMatrix(FERelPers.row, 5) = "0.00"
    End If
    
    If Not ValidaDatosPersRelacion Then
        Exit Sub
    End If
    
   If cmdPersRelaEjecutado = 1 Then
        Call oPersona.AdicionaPersonaRelacion
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaNueva, FERelPers.row - 1)
    Else
        If cmdPersRelaEjecutado = 2 Then
            If oPersona.ObtenerRelaPersTipoAct(FERelPers.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRelacPersTipoAct(PersFilaModificada, FERelPers.row - 1)
            End If
        End If
    End If
    
    'Apellidos y Nombres
    Call oPersona.ActualizaPersRelaPersona(FERelPers.TextMatrix(FERelPers.row, 2), FERelPers.TextMatrix(FERelPers.row, 1), FERelPers.row - 1)
    ' Relacion
    Call oPersona.ActualizaPersRelaRelacion(FERelPers.TextMatrix(FERelPers.row, 3), FERelPers.row - 1)
    'Beneficiario
    Call oPersona.ActualizarRelaPersBenef(FERelPers.TextMatrix(FERelPers.row, 4), FERelPers.row - 1)
    'Beneficiario Porcentaje
    Call oPersona.ActualizarRelaPersBenefPorc(CDbl(FERelPers.TextMatrix(FERelPers.row, 5)), FERelPers.row - 1)
    'AMP
    Call oPersona.ActualizarRelaPersAMP(FERelPers.TextMatrix(FERelPers.row, 6), FERelPers.row - 1)
    
    'Habilitar Controles
    cmdPersRelaEjecutado = 0
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FERelPers.SetFocus
End Sub

Private Sub cmdPersRelacCancelar_Click()
    CargaRelacionesPersonas
    'Habilitar Controles
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    cmdPersRelacAceptar.Visible = False
    cmdPersRelacCancelar.Visible = False
    
    FERelPers.SetFocus

End Sub

Private Sub cmdPersRelacDel_Click()
    If MsgBox("Esta Seguro que Desea Eliminar La Relacion con esta Persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaEliminda, FERelPers.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRelacionesPersonas
    End If
End Sub

Private Sub cmdPersRelacEditar_Click()
    If oPersona.NumeroRelacPers > 0 Then
        cmdPersRelaEjecutado = 2
        FERelPersNoMoverdeFila = FERelPers.row
        NomMoverSSTabs = SSTabs.Tab
        FERelPers.lbEditarFlex = True
        FERelPers.SetFocus
        cmdPersRelacNew.Enabled = False
        cmdPersRelacEditar.Enabled = False
        cmdPersRelacDel.Enabled = False
        cmdPersRelacAceptar.Visible = True
        cmdPersRelacCancelar.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        FERelPers.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdPersRelacNew_Click()
    cmdPersRelacAceptar.Visible = True
    cmdPersRelacCancelar.Visible = True
    cmdPersRelacNew.Enabled = False
    cmdPersRelacDel.Enabled = False
    cmdPersRelacEditar.Enabled = False
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    cmdPersRelaEjecutado = 1
    FERelPersNoMoverdeFila = FERelPers.Rows - 1
    FERelPers.SetFocus
End Sub

Private Sub cmdRefBanAcepta_Click()

    If Not ValidaDatosRefBancaria Then
        Exit Sub
    End If
        
   If cmdPersRefBancariaEjecutado = 1 Then
        Call oPersona.AdicionaRefBancaria
        Call oPersona.ActualizarRefBanTipoAct(PersFilaNueva, feRefBancaria.row - 1)
    Else
        If cmdPersRefBancariaEjecutado = 2 Then
            If oPersona.ObtenerRefBanTipoAct(feRefBancaria.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRefBanTipoAct(PersFilaModificada, feRefBancaria.row - 1)
            End If
        End If
    End If
    
    'Codigo de la Institucion Financiera
    Call oPersona.ActualizaRefBanCodIF(feRefBancaria.TextMatrix(feRefBancaria.row, 1), feRefBancaria.row - 1)
    'Nombre de la Institución Financiera
    Call oPersona.ActualizaRefBanNombre(feRefBancaria.TextMatrix(feRefBancaria.row, 2), feRefBancaria.row - 1)
    'Número de Cuenta de
    Call oPersona.ActualizaRefBanNumCta(feRefBancaria.TextMatrix(feRefBancaria.row, 3), feRefBancaria.row - 1)
    'Número de Tarjeta
    Call oPersona.ActualizaRefBanNumTar(feRefBancaria.TextMatrix(feRefBancaria.row, 4), feRefBancaria.row - 1)
    'Línea de Crédito
    Call oPersona.ActualizaRefBanLinCred(feRefBancaria.TextMatrix(feRefBancaria.row, 5), feRefBancaria.row - 1)
                
    'Habilitar Controles
    cmdPersRefBancariaEjecutado = 0
    FERefBanPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefBancaria.lbEditarFlex = False
    cmdRefBanNuevo.Enabled = True
    cmdRefBanEdita.Enabled = True
    cmdRefBanElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdRefBanAcepta.Visible = False
    cmdRefBanCancela.Visible = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    feRefBancaria.SetFocus

End Sub

Private Sub cmdRefBanCancela_Click()
    CargaRefBancarias
    'Habilitar Controles
    FERefBanPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefBancaria.lbEditarFlex = False
    cmdRefBanNuevo.Enabled = True
    cmdRefBanEdita.Enabled = True
    cmdRefBanElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdRefBanAcepta.Visible = False
    cmdRefBanCancela.Visible = False
    
    feRefBancaria.SetFocus

End Sub

Private Sub cmdRefBanEdita_Click()
   If oPersona.NumeroRefBancaria > 0 Then
        cmdPersRefBancariaEjecutado = 2
        FERefBanPersNoMoverdeFila = feRefBancaria.row
        NomMoverSSTabs = SSTabs.Tab
        feRefBancaria.lbEditarFlex = True
        feRefBancaria.SetFocus
        cmdRefBanNuevo.Enabled = False
        cmdRefBanEdita.Enabled = False
        cmdRefBanElimina.Enabled = False
        cmdRefBanAcepta.Visible = True
        cmdRefBanCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        feRefBancaria.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdRefBanElimina_Click()
    If MsgBox("Esta Seguro que Desea Eliminar la Referencia Bancaria", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRefBanTipoAct(PersFilaEliminda, feRefBancaria.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRefBancarias
    End If
End Sub

Private Sub cmdRefBanNuevo_Click()
    cmdRefBanAcepta.Visible = True
    cmdRefBanCancela.Visible = True
    cmdRefBanNuevo.Enabled = False
    cmdRefBanElimina.Enabled = False
    cmdRefBanEdita.Enabled = False
    feRefBancaria.lbEditarFlex = True
    feRefBancaria.AdicionaFila
    cmdPersRefBancariaEjecutado = 1
    FERefBanPersNoMoverdeFila = feRefBancaria.Rows - 1
    feRefBancaria.SetFocus
End Sub

Private Sub cmdRefComAcepta_Click()
    
    If Not ValidaDatosRefComercial Then
        Exit Sub
    End If
    
   If cmdPersRefComercialEjecutado = 1 Then
        Call oPersona.AdicionaRefComercial
        Call oPersona.ActualizarRefComTipoAct(PersFilaNueva, feRefComercial.row - 1)
        lnNumRefCom = lnNumRefCom + 1
        feRefComercial.TextMatrix(feRefComercial.row, 4) = lnNumRefCom
    Else
        If cmdPersRefComercialEjecutado = 2 Then
            If oPersona.ObtenerRefComTipoAct(feRefComercial.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRefComTipoAct(PersFilaModificada, feRefComercial.row - 1)
            End If
        End If
    End If
    
    'Nombre/Razón Social
    Call oPersona.ActualizaRefComNombre(feRefComercial.TextMatrix(feRefComercial.row, 1), feRefComercial.row - 1)
    'Tipo de Referencia Comercial
    Call oPersona.ActualizaRefComTipoRel(feRefComercial.TextMatrix(feRefComercial.row, 2), feRefComercial.row - 1)
    'Telefono
    Call oPersona.ActualizaRefComFono(feRefComercial.TextMatrix(feRefComercial.row, 3), feRefComercial.row - 1)
    Call oPersona.ActualizaRefComCod(feRefComercial.TextMatrix(feRefComercial.row, 4), feRefComercial.row - 1)
            
    'Habilitar Controles
    cmdPersRefComercialEjecutado = 0
    FERefComPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefComercial.lbEditarFlex = False
    cmdRefComNuevo.Enabled = True
    cmdRefComEdita.Enabled = True
    cmdRefComElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdRefComAcepta.Visible = False
    cmdRefComCancela.Visible = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    feRefComercial.SetFocus

End Sub

Private Sub cmdRefComCancela_Click()
    
    CargaRefComerciales
    'Habilitar Controles
    FERefComPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefComercial.lbEditarFlex = False
    cmdRefComNuevo.Enabled = True
    cmdRefComEdita.Enabled = True
    cmdRefComElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdRefComAcepta.Visible = False
    cmdRefComCancela.Visible = False
    
    feRefComercial.SetFocus
    
End Sub

Private Sub cmdRefComEdita_Click()
    If oPersona.NumeroRefComercial > 0 Then
        cmdPersRefComercialEjecutado = 2
        FERefComPersNoMoverdeFila = feRefComercial.row
        NomMoverSSTabs = SSTabs.Tab
        feRefComercial.lbEditarFlex = True
        feRefComercial.SetFocus
        cmdRefComNuevo.Enabled = False
        cmdRefComEdita.Enabled = False
        cmdRefComElimina.Enabled = False
        cmdRefComAcepta.Visible = True
        cmdRefComCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        feRefComercial.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdRefComElimina_Click()
    If MsgBox("Esta Seguro que Desea Eliminar la Referencia Comercial", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRefComTipoAct(PersFilaEliminda, feRefComercial.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRefComerciales
    End If
End Sub

Private Sub cmdRefComNuevo_Click()
    cmdRefComAcepta.Visible = True
    cmdRefComCancela.Visible = True
    cmdRefComNuevo.Enabled = False
    cmdRefComElimina.Enabled = False
    cmdRefComEdita.Enabled = False
    feRefComercial.lbEditarFlex = True
    feRefComercial.AdicionaFila
    cmdPersRefComercialEjecutado = 1
    FERefComPersNoMoverdeFila = feRefComercial.Rows - 1
    feRefComercial.SetFocus
End Sub

Private Sub cmdSalir_Click()
    If bPersonaAct Then
        MsgBox "Grabe o Cancele los Cambios Antes de Salir", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oPersona = Nothing
    Unload Me
End Sub


Private Sub FEDocs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If FEDocs.col = 2 Then
            If cmdPersIDAceptar.Enabled Then
                cmdPersIDAceptar.SetFocus
            End If
        End If
    End If
End Sub

Private Sub FEDocs_RowColChange()
    If cmdPersDocEjecutado = 1 Or cmdPersDocEjecutado = 2 Then
        FEDocs.row = FEDocsPersNoMoverdeFila
    End If
End Sub

Private Sub fePatVehicular_RowColChange()
    
    If fePatVehicular.col = 2 Then
        If fePatVehicular.TextMatrix(fePatVehicular.row, 2) <> "" Then
            If CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) < 1000 Or CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) > Year(gdFecSis) Then
                MsgBox "Año de Fabricación no válido", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
    
    If cmdPersPatVehicularEjecutado = 1 Or cmdPersPatVehicularEjecutado = 2 Then
        fePatVehicular.row = FEPatVehPersNoMoverdeFila
    End If

End Sub

Private Sub feRefBancaria_OnCellChange(pnRow As Long, pnCol As Long)
    If feRefBancaria.col = 5 Then
        If feRefBancaria.TextMatrix(feRefBancaria.row, 5) <> "" Then
            If CCur(feRefBancaria.TextMatrix(feRefBancaria.row, 5)) < 0 Then
                MsgBox "Linea de Crédito no puede ser negativa", vbInformation, "Aviso"
                feRefBancaria.TextMatrix(feRefBancaria.row, 5) = 0
            End If
        End If
    End If
End Sub

Private Sub feRefBancaria_RowColChange()
    
    If cmdPersRefBancariaEjecutado = 1 Or cmdPersRefBancariaEjecutado = 2 Then
        feRefBancaria.row = FERefBanPersNoMoverdeFila
    End If
End Sub

Private Sub feRefComercial_RowColChange()
    
    If cmdPersRefComercialEjecutado = 1 Or cmdPersRefComercialEjecutado = 2 Then
        feRefComercial.row = FERefComPersNoMoverdeFila
    End If
End Sub

Private Sub FERelPers_Click()
    Call FERelPers_RowColChange
End Sub


Private Sub FERelPers_EnterCell()
    FERelPers_RowColChange
    
End Sub

Private Sub FERelPers_RowColChange()
Dim oConstante As DConstante
    If FERelPers.lbEditarFlex Then
        If FERelPersNoMoverdeFila <> -1 Then
        
            If TxtBCodPers.psCodigoPersona = FERelPers.TextMatrix(FERelPers.row, 1) Then
                FERelPers.TextMatrix(FERelPers.row, 1) = ""
                FERelPers.TextMatrix(FERelPers.row, 2) = ""
                Exit Sub
            End If
        
            FERelPers.row = FERelPersNoMoverdeFila
        End If
        Set oConstante = New DConstante
        Select Case FERelPers.col
            Case 3 'Relacion de Persona
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
            Case 4 'Beneficiario
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacBenef)
            Case 6 'AMP
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacAmp)
        End Select
        Set oConstante = Nothing
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Screen.MousePointer = 11
    Call CargaControles
    TxtBCodPers.Enabled = True
    cmdEditar.Enabled = False
    Screen.MousePointer = 0
    bNuevaPersona = False
    'WIOR 20130827 ***************************
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If bPersonaAct Then
        MsgBox "Grabe o Cancele los Cambios Antes de Salir", vbInformation, "Aviso"
        Cancel = 1
    End If
End Sub

Private Sub SSTabs_Click(PreviousTab As Integer)
    If NomMoverSSTabs > -1 Then
        SSTabs.Tab = NomMoverSSTabs
    End If
End Sub

Private Sub SSTIdent_LostFocus()
    If cmdPersIDAceptar.Visible Then
        MsgBox "Acepte o Cancele el Ingreso del Documento", vbInformation, "Aviso"
        cmdPersIDAceptar.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtApellidoCasada_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoCasada = Trim(TxtApellidoCasada.Text)
    End If
End Sub

Private Sub TxtApellidoCasada_GotFocus()
    fEnfoque TxtApellidoCasada
End Sub

Private Sub TxtApellidoCasada_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtPersNombreN.SetFocus
    End If
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New DPersona
    oPersona.sCodAge = gsCodAge
    Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
    If oPersona.PersCodigo = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    Call CargaDatos
    If SSTabs.Enabled And SSTabs.Visible Then
        SSTabs.SetFocus
    End If
End Sub

Private Sub LimpiarPantalla()
Dim I As Integer
    
    bEstadoCargando = True
    'TxtBCodPers.Text = ""
    cmbPersPersoneria.ListIndex = -1
    txtPersNombreAP.Text = ""
    txtPersNombreAM.Text = ""
    txtPersNombreN.Text = ""
    cmbPersNatSexo.ListIndex = -1
    cmbPersNatEstCiv.ListIndex = -1
    txtPersNatHijos.Text = "0"
    txtPersNacCreac.Text = "__/__/____"
    txtPersTelefono.Text = ""
    CboPersCiiu.ListIndex = -1
    cmbPersEstado.ListIndex = -1
    TxtTalla.Text = "0.00"
    TxtPeso.Text = "0.00"
    CboTipoSangre.ListIndex = -1
    chkResidente.value = 0
    txtPersTelefono2.Text = ""
    'EJVG20111212 ****************
    txtCel1.Text = ""
    txtCel2.Text = ""
    txtCel3.Text = ""
    
    TxtEmail.Text = ""
    TxtEmail2.Text = "" 'EJVG20111212
    Call LimpiaFlex(FEDocs)
    
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), "04028")
    cmbPersUbiGeo(1).ListIndex = -1
    cmbPersUbiGeo(2).ListIndex = -1
    cmbPersUbiGeo(3).ListIndex = -1
    cmbPersUbiGeo(4).ListIndex = -1
    txtPersDireccDomicilio.Text = ""
    cmbPersDireccCondicion.ListIndex = -1
    txtValComercial.Text = ""
    txtPersNombreRS.Text = ""
    TxtSiglas.Text = ""
    cmbPersJurTpo.ListIndex = -1
    cmbPersJurMagnitud.ListIndex = -1
    txtPersJurEmpleados.Text = ""
    TxtSbs.Text = ""
    CmbRela.ListIndex = -1
    Call LimpiaFlex(FERelPers)
    Call LimpiaFlex(FEFteIng)
    Call LimpiaFlex(feRefComercial)
    IDBFirma.RutaImagen = ""
    bEstadoCargando = False
End Sub

Private Sub TxtBCodPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If cmbPersPersoneria.Enabled Then
            cmbPersPersoneria.SetFocus
        Else
            SSTabs.Tab = 0
            If txtPersNombreAP.Enabled Then
                txtPersNombreAP.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtCel1_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular = Trim(txtCel1.Text)
    End If
End Sub

Private Sub txtCel1_KeyPress(KeyAscii As Integer)
    If txtCel1.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                txtCel2.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            txtCel2.SetFocus
        End If
    End If
End Sub

Private Sub txtCel2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular2 = Trim(txtCel2.Text)
    End If
End Sub

Private Sub txtCel2_KeyPress(KeyAscii As Integer)
    If txtCel2.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                txtCel3.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            txtCel3.SetFocus
        End If
    End If
End Sub

Private Sub txtCel3_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular3 = Trim(txtCel3.Text)
    End If
End Sub

Private Sub txtCel3_KeyPress(KeyAscii As Integer)
    If txtCel2.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                TxtEmail.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            TxtEmail.SetFocus
        End If
    End If
End Sub

Private Sub TxtEmail_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Email = Trim(TxtEmail.Text)
    End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtEmail2.SetFocus
    End If
End Sub

Private Sub TxtEmail2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Email2 = Trim(TxtEmail2.Text)
    End If
End Sub

Private Sub TxtEmail2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboPersCiiu.SetFocus
    End If
End Sub

Private Sub txtPersDireccDomicilio_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Domicilio = Trim(txtPersDireccDomicilio.Text)
    End If
End Sub

Private Sub txtPersDireccDomicilio_GotFocus()
    fEnfoque txtPersDireccDomicilio
End Sub

Private Sub txtPersDireccDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersDireccCondicion.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumerosEmpleados = CInt(IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text)))
    End If
End Sub

Private Sub txtPersJurEmpleados_GotFocus()
    fEnfoque txtPersJurEmpleados
End Sub

Private Sub txtPersJurEmpleados_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersJurMagnitud.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_LostFocus()
    txtPersJurEmpleados.Text = IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text))
End Sub

Private Sub txtPersNacCreac_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersNacCreac.Text))) = 0 Then
            oPersona.FechaNacimiento = CDate(txtPersNacCreac.Text)
        End If
    End If
    
End Sub

Private Sub txtPersNacCreac_GotFocus()
    fEnfoque txtPersNacCreac
End Sub

Private Sub txtPersNacCreac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtSbs.SetFocus
    End If
End Sub

Private Sub txtPersNacCreac_LostFocus()
Dim sCad As String

    sCad = ValidaFecha(txtPersNacCreac.Text)
    '***Modificado por ELRO el 20111013, según Acta 270-2011/TI-D
    'If Not Trim(sCad) = "" Then
    '    MsgBox sCad, vbInformation, "Aviso"
    '    If txtPersNacCreac.Enabled Then txtPersNacCreac.SetFocus
    '    Exit Sub
    'End If
    'If CDate(txtPersNacCreac.Text) >= gdFecSis Then
    '     MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
    '     txtPersNacCreac.SetFocus
    '     Exit Sub
    'End If
    
    If gnPersPersoneria = 1 And gsCodArea <> "022" Then
       If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            If txtPersNacCreac.Enabled Then txtPersNacCreac.SetFocus
            Exit Sub
        End If
        
        If CDate(txtPersNacCreac.Text) >= gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtPersNacCreac.SetFocus
            Exit Sub
        End If
    End If
        
    
    'Fin Modificado por ELRO*************************************
    
    

End Sub

Private Sub txtPersNatHijos_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Hijos = CInt(Trim(IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)))
      End If
End Sub

Private Sub txtPersNatHijos_GotFocus()
    fEnfoque txtPersNatHijos
End Sub

Private Sub txtPersNatHijos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbNacionalidad.SetFocus
    End If
End Sub

Private Sub txtPersNatHijos_LostFocus()
    txtPersNatHijos.Text = IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)
End Sub

Private Sub txtPersNombreAM_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.ApellidoMaterno = Trim(txtPersNombreAM.Text)
   End If
End Sub

Private Sub txtPersNombreAM_GotFocus()
    fEnfoque txtPersNombreAM
End Sub

Private Sub txtPersNombreAM_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If TxtApellidoCasada.Visible Then
            TxtApellidoCasada.SetFocus
        Else
            txtPersNombreN.SetFocus
        End If
    End If
End Sub

Private Sub txtPersNombreAP_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoPaterno = Trim(txtPersNombreAP.Text)
    End If
End Sub

Private Sub txtPersNombreAP_GotFocus()
    fEnfoque txtPersNombreAP
End Sub

Private Sub txtPersNombreAP_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtPersNombreAM.SetFocus
    End If
End Sub

Private Sub txtPersNombreN_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
      oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.nombres = Trim(txtPersNombreN.Text)
  End If
End Sub

Private Sub txtPersNombreN_GotFocus()
    fEnfoque txtPersNombreN
End Sub

Private Sub txtPersNombreN_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersNatSexo.SetFocus
    End If
End Sub

Private Sub txtPersNombreRS_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NombreCompleto = Trim(txtPersNombreRS.Text)
    End If
End Sub

Private Sub txtPersNombreRS_GotFocus()
    fEnfoque txtPersNombreRS
End Sub

Private Sub txtPersNombreRS_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtSiglas.SetFocus
    End If
End Sub

Private Sub txtPersTelefono_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos = Trim(txtPersTelefono.Text)
    End If
End Sub

Private Sub txtPersTelefono_GotFocus()
    fEnfoque txtPersTelefono
End Sub

Private Sub txtPersTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtPersTelefono2.SetFocus
    End If
End Sub

Private Sub txtPersTelefono2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos2 = Trim(txtPersTelefono2.Text)
    End If
End Sub

Private Sub txtPersTelefono2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        'TxtEmail.SetFocus
        txtCel1.SetFocus
    End If
End Sub

Private Sub TxtPeso_Change()
    On Error Resume Next
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Peso = CDbl(Format(IIf(Trim(TxtPeso.Text) = "", "0", TxtPeso.Text), "#0.00"))
      End If
End Sub

Private Sub TxtPeso_GotFocus()
    fEnfoque TxtPeso
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPeso, KeyAscii, 5, 2)
    If KeyAscii = 13 Then
        If bPermisoEditarTodo Then 'EJVG20111219
            txtPersNacCreac.SetFocus
        Else
            TxtSbs.SetFocus
        End If
    End If
End Sub

Private Sub TxtPeso_LostFocus()
    If Trim(TxtPeso.Text) = "." Then
        TxtPeso.Text = "0.00"
    End If
    TxtPeso.Text = Format(IIf(Trim(TxtPeso.Text) = "", "0.00", TxtPeso.Text), "#0.00")
End Sub

Private Sub TxtSbs_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PersCodSbs = Trim(TxtSbs.Text)
    End If
End Sub

Private Sub TxtSbs_GotFocus()
    fEnfoque TxtSbs
End Sub

Private Sub TxtSbs_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
         txtPersTelefono.SetFocus
    End If
End Sub

Private Sub TxtSiglas_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Siglas = Trim(TxtSiglas.Text)
    End If
End Sub

Private Sub txtSiglas_GotFocus()
    fEnfoque TxtSiglas
End Sub

Private Sub txtSiglas_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersJurTpo.SetFocus
    End If
End Sub

Private Sub TxtTalla_Change()
    On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Talla = CDbl(IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text)))
    End If
End Sub

Private Sub TxtTalla_GotFocus()
    fEnfoque TxtTalla
End Sub

Private Sub TxtTalla_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtTalla, KeyAscii, 4, 2)
    If KeyAscii = 13 Then
        CboTipoSangre.SetFocus
    End If
End Sub

Private Sub TxtTalla_LostFocus()
    If TxtTalla.Text = "." Then
        TxtTalla.Text = "0.00"
    End If
    TxtTalla.Text = Format(IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text)), "#0.00")
End Sub

Private Sub txtValComercial_Change()
    On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ValComDomicilio = txtValComercial.Text
    End If
End Sub

Private Sub txtValComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPersIDnew.SetFocus
    End If
End Sub

'EJVG20111219 ***************************************************
Public Function ObtenerVecesCreditoyAhorroPersona(ByVal psPersCod As String) As Integer
    Dim oCred As DCredito
    Set oCred = New DCredito
    ObtenerVecesCreditoyAhorroPersona = oCred.ObtenerVecesCreditoyAhorroPersona(psPersCod)
    Set oCred = Nothing
End Function
Public Function validaPermisoEditarPersona(ByVal psCargo As String) As Boolean
    Dim oPersona As New DPersona
    validaPermisoEditarPersona = oPersona.validaPermisoEditarPersona(psCargo)
    Set oPersona = Nothing
End Function

