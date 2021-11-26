VERSION 5.00
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "SICMACT Sistema Integrado de la Caja Municipal de Ahorro y Credito de Trujillo"
   ClientHeight    =   7035
   ClientLeft      =   750
   ClientTop       =   1380
   ClientWidth     =   11880
   Icon            =   "MDISicmact1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Tiempo 
      Interval        =   60000
      Left            =   1485
      Top             =   1365
   End
   Begin VB.PictureBox SBBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   6780
      Width           =   11880
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
   End
   Begin VB.Menu M0100000000 
      Caption         =   "Arc&hivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   1
      End
   End
   Begin VB.Menu M0200000000 
      Caption         =   "&Creditos"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Definiciones"
         Index           =   0
         Begin VB.Menu M0201010000 
            Caption         =   "&Parametros de Control"
            Index           =   0
            Begin VB.Menu M0201010100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Lineas de Credito"
            Index           =   1
            Begin VB.Menu M0201010200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0201010200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0201010200 
               Caption         =   "&Consulta"
               Index           =   3
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Niveles de Aprobacion"
            Index           =   2
            Begin VB.Menu M0201010300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010300 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Gastos"
            Index           =   3
            Begin VB.Menu M0201010400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010400 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Solicitud Credito"
         Index           =   1
         Begin VB.Menu M0201020000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Relaciones de Credito"
         Index           =   2
         Begin VB.Menu M0201030000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201030000 
            Caption         =   "Reasignacion de &Cartera en Lote"
            Index           =   1
         End
         Begin VB.Menu M0201030000 
            Caption         =   "Con&sulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Garantias"
         Index           =   3
         Begin VB.Menu M0201040000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M0201040000 
            Caption         =   "Gravamen"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Sugerencia"
         Index           =   4
         Begin VB.Menu M0201050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Resolver Creditos"
         Index           =   5
         Begin VB.Menu M0201060000 
            Caption         =   "&Aprobacion"
            Index           =   0
         End
         Begin VB.Menu M0201060000 
            Caption         =   "&Rechazo"
            Index           =   1
         End
         Begin VB.Menu M0201060000 
            Caption         =   "A&nulacion Creditos Aprobados"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reprogramacion Credito"
         Index           =   6
         Begin VB.Menu M0201070000 
            Caption         =   "Repr&ogramacion"
            Index           =   0
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Reprogramacion en &Lote"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Refinanciacion"
         Index           =   7
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Actualizacion de &Metodos de Liquidacion"
         Index           =   8
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Perdonar Mora"
         Index           =   9
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Gastos"
         Index           =   10
         Begin VB.Menu M0201080000 
            Caption         =   "&Administracion de Gastos en &Lote"
            Index           =   0
         End
         Begin VB.Menu M0201080000 
            Caption         =   "Mantenimiento de Penalidad de Cancelacion"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Reasignar &Institucion"
         Index           =   11
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Transferencia a Recuperaciones"
         Index           =   12
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Analista"
         Index           =   13
         Begin VB.Menu M0201090000 
            Caption         =   "&Nota de Analista"
            Index           =   0
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&Metas de Analista"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Si&mulaciones"
         Index           =   14
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de &Pagos"
            Index           =   0
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de &Desembolsos Parciales"
            Index           =   1
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de Cuota &Libre"
            Index           =   2
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Simulador de Pagos"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Cons&ultas"
         Index           =   15
         Begin VB.Menu M0201200000 
            Caption         =   "&Historial de Credito"
            Index           =   0
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reportes"
         Index           =   16
         Begin VB.Menu M0201300000 
            Caption         =   "&Duplicados"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "&Pignoraticio"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "&Contratro"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Mantenimiento Descripcion"
            Index           =   1
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Anulación"
            Index           =   2
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Bloqueo"
            Index           =   3
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Rescate Joyas"
         Index           =   2
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Remate"
         Index           =   3
         Begin VB.Menu M0301020000 
            Caption         =   "Preparacion Remate"
            Index           =   0
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Remate"
            Index           =   1
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Adjudicacion"
         Index           =   4
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Subasta"
         Index           =   5
         Begin VB.Menu M0301030000 
            Caption         =   "Preparacion Subasta"
            Index           =   0
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Subasta"
            Index           =   1
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Reportes"
         Index           =   6
         Begin VB.Menu M0301040000 
            Caption         =   "Movimientos Diario"
            Index           =   0
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Listados Generales"
            Index           =   1
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Estadisticas"
            Index           =   2
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Contrato de Persona"
            Index           =   3
         End
      End
   End
   Begin VB.Menu M0400000000 
      Caption         =   "&Captaciones"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "&Parámetros"
         Index           =   0
         Begin VB.Menu M0401010000 
            Caption         =   "&Consulta"
            Index           =   0
         End
         Begin VB.Menu M0401010000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Tasas de Interés"
         Index           =   1
         Begin VB.Menu M0401020000 
            Caption         =   "&Ahorros"
            Index           =   0
            Begin VB.Menu M0401020100 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
         Begin VB.Menu M0401020000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
            Begin VB.Menu M0401020200 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
         Begin VB.Menu M0401020000 
            Caption         =   "&CTS"
            Index           =   2
            Begin VB.Menu M0401020300 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Mantenimiento"
         Index           =   2
         Begin VB.Menu M0401030000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0401030000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0401030000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Bloqueo/Desbloqueo"
         Index           =   3
         Begin VB.Menu M0401040000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0401040000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0401040000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Simulación Plazo Fijo"
         Index           =   4
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Tarjeta Magnética"
         Index           =   5
         Begin VB.Menu M0401050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Re&lación"
            Index           =   1
         End
         Begin VB.Menu M0401050000 
            Caption         =   "&Bloqueo/Desbloqueo"
            Index           =   2
         End
         Begin VB.Menu M0401050000 
            Caption         =   "&Cambio de Clave"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Bene&ficiarios"
         Index           =   6
         Begin VB.Menu M0401060000 
            Caption         =   "&Consulta"
            Index           =   0
         End
         Begin VB.Menu M0401060000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Orden Pago"
         Index           =   7
         Begin VB.Menu M0401070000 
            Caption         =   "&Generación y Emisión"
            Index           =   0
         End
         Begin VB.Menu M0401070000 
            Caption         =   "&Certificación"
            Index           =   1
         End
         Begin VB.Menu M0401070000 
            Caption         =   "&Anulación"
            Index           =   2
         End
         Begin VB.Menu M0401070000 
            Caption         =   "Con&sulta"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Ser&vicios"
         Index           =   8
         Begin VB.Menu M0401080000 
            Caption         =   "&Sedalib"
            Index           =   0
            Begin VB.Menu M0401080100 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080100 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0401080100 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Hidrandina"
            Index           =   1
            Begin VB.Menu M0401080200 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080200 
               Caption         =   "&Reportes"
               Index           =   1
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Universidad Nacional de Trujillo"
            Index           =   2
            Begin VB.Menu M0401080300 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080300 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0401080300 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Instituciones Convenio"
            Index           =   3
            Begin VB.Menu M0401080400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0401080400 
               Caption         =   "&Cuentas"
               Index           =   1
            End
            Begin VB.Menu M0401080400 
               Caption         =   "&Plan de Pagos"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu M0500000000 
      Caption         =   "&Recuperaciones"
      Index           =   0
      Begin VB.Menu M0501000000 
         Caption         =   "&Comisiones Abogados"
         Index           =   0
      End
      Begin VB.Menu M0501000000 
         Caption         =   "&Registro de Otros Creditos "
         Index           =   1
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Proceso &Judicial"
         Index           =   2
         Begin VB.Menu M0501010000 
            Caption         =   "&Expedientes Judiciales"
            Index           =   0
         End
         Begin VB.Menu M0501010000 
            Caption         =   "Actuaciones &Procesales"
            Index           =   1
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "&Gastos de Recuperacines"
         Index           =   3
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Metodo de &Liquidacion"
         Index           =   4
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Cas&tigar Credito"
         Index           =   5
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Pago de Creditos"
         Index           =   6
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Extorno de Pagos"
         Index           =   7
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Consulta Para&metros"
         Index           =   8
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Mantenimiento Parametros"
         Index           =   9
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "&Operaciones"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "Desembolso &Abono a Cuenta"
         Index           =   0
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones CMACs &Recepción"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones CMACs &Llamada"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "&Extornos"
         Index           =   4
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Credito"
            Index           =   0
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Pignoraticio"
            Index           =   1
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Recuperaciones"
            Index           =   2
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Captaciones"
            Index           =   3
         End
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "Perso&nas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Manteniemiento"
            Index           =   1
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Instituciones Financieras"
         Index           =   1
         Begin VB.Menu M0701020000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0701020000 
            Caption         =   "Consulta"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "Cli&ente"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "&Posicion de Cliente"
         Index           =   0
      End
   End
   Begin VB.Menu M0900000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0901000000 
         Caption         =   "&Spooler de Impresion"
         Index           =   0
      End
      Begin VB.Menu M0901000000 
         Caption         =   "Configuracion &Perifericos"
         Index           =   1
      End
   End
   Begin VB.Menu M1000000000 
      Caption         =   "&Seguridad"
      Index           =   0
      Begin VB.Menu M1001000000 
         Caption         =   "Permisos"
         Index           =   0
      End
      Begin VB.Menu M1001000000 
         Caption         =   "Prueba3"
         Index           =   1
      End
   End
   Begin VB.Menu M1600000000 
      Caption         =   "&Recursos Humanos"
      Begin VB.Menu M1601000000 
         Caption         =   "&Selección"
         Index           =   0
         Begin VB.Menu M1601010000 
            Caption         =   "&Inicio Proceso Selección"
            Index           =   0
            Begin VB.Menu M1601010100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Postulantes"
            Index           =   1
            Begin VB.Menu M1601010200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Evaluación"
            Index           =   2
            Begin VB.Menu M1601010300 
               Caption         =   "&Curricular"
               Index           =   0
               Begin VB.Menu M1601010301 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "&Escrita"
               Index           =   1
               Begin VB.Menu M1601010302 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "&Psicologica"
               Index           =   2
               Begin VB.Menu M1601010303 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "E&ntrevista"
               Index           =   3
               Begin VB.Menu M1601010304 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Resultados y Cierre"
            Index           =   3
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Consulta"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Contratos"
         Index           =   2
         Begin VB.Menu M1601020000 
            Caption         =   "&Registro basado en el Proceso de Selección"
            Index           =   0
         End
         Begin VB.Menu M1601020000 
            Caption         =   "R&egistro manual"
            Index           =   1
         End
         Begin VB.Menu M1601020000 
            Caption         =   "&Mantenimiento"
            Index           =   2
         End
         Begin VB.Menu M1601020000 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu M1601020000 
            Caption         =   "Re&scindir"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Adenda"
         Index           =   4
         Begin VB.Menu M1601030000 
            Caption         =   "&Registro"
            Index           =   0
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu M1601000000 
         Caption         =   "C&urriculum Vitae"
         Index           =   6
         Begin VB.Menu M1601040000 
            Caption         =   "&Configuracion"
            Index           =   0
            Begin VB.Menu M1601040100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601040100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Registro"
            Index           =   1
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Mantenimiento"
            Index           =   2
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Reporte"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Ho&rario Laboral"
         Index           =   8
         Begin VB.Menu M1601050000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601050000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Asistencia"
         Index           =   9
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu M1601000000 
         Caption         =   "E&valuacion Interna"
         Index           =   11
         Begin VB.Menu M1601060000 
            Caption         =   "&Inicio de proceso de Evaluacion Interna"
            Index           =   0
            Begin VB.Menu M1601060100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060100 
               Caption         =   "&Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Curricular"
            Index           =   1
            Begin VB.Menu M1601060200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Escrita"
            Index           =   2
            Begin VB.Menu M1601060300 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Psicologica"
            Index           =   3
            Begin VB.Menu M1601060400 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "E&ntrevista"
            Index           =   4
            Begin VB.Menu M1601060500 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Registro"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "Resultados y Cie&rre"
            Index           =   5
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Consulta"
            Index           =   6
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Permisos"
         Index           =   13
         Begin VB.Menu M1601070000 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu M1601070100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601070100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601070000 
            Caption         =   "&Autorizacion/Rechazo"
            Index           =   1
         End
         Begin VB.Menu M1601070000 
            Caption         =   "&Reporte"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Vacaciones"
         Index           =   14
         Begin VB.Menu M1601080000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601080000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Descansos"
         Index           =   15
         Begin VB.Menu M1601090000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601090000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Suspensiones"
         Index           =   16
         Begin VB.Menu M1601100000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601100000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Meritos y Demeritos"
         Index           =   18
         Begin VB.Menu M1601110000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601110100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601110100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601110000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M1601110000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Car&gos Laborales"
         Index           =   19
         Begin VB.Menu M1601120000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601120100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601120100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601120000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Sueldos"
         Index           =   20
         Begin VB.Menu M1601130000 
            Caption         =   "&Registro"
            Index           =   0
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Sistema de Pensiones"
         Index           =   21
         Begin VB.Menu M1601140000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601140100 
               Caption         =   "&Mantenimiento"
            End
         End
         Begin VB.Menu M1601140000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "InformeSocial"
         Index           =   22
         Begin VB.Menu M1601150000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Reportes"
            Index           =   3
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Comentario"
         Index           =   23
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Asistencia Medica"
         Index           =   24
         Begin VB.Menu M1601160000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601160100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
         End
         Begin VB.Menu M1601160000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Confi&guracion de Planilla de Remuneraciones"
         Index           =   26
         Begin VB.Menu M1601170000 
            Caption         =   "&Configuracion de Conceptos Remunerativos"
            Index           =   0
            Begin VB.Menu M1601170100 
               Caption         =   "&Mantenimiento de Conceptos Remunerativos"
               Index           =   0
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Mantenimiento de &Tabla Alias"
               Index           =   1
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601170000 
            Caption         =   "C&onfiguracion de Conceptos de Planilla de Remuneraciones"
            Index           =   1
            Begin VB.Menu M1601170200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Planilla de remuneraciones del RR.HH."
         Index           =   27
         Begin VB.Menu M1601180000 
            Caption         =   "&Conceptos Fijos de Planilla de Remuneraciones"
            Index           =   0
            Begin VB.Menu M1601180100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180100 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M1601180100 
               Caption         =   "&Reporte"
               Index           =   2
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "Conceptos &Variables de Planilla de Remuneraciones"
            Index           =   1
            Begin VB.Menu M1601180200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180200 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M1601180200 
               Caption         =   "&Reporte"
               Index           =   2
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "E&xtra Planilla Variables (Ingresos y Descuentos)"
            Index           =   2
            Begin VB.Menu M1601180300 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   28
      End
      Begin VB.Menu M1601000000 
         Caption         =   "P&rocesos"
         Index           =   29
         Begin VB.Menu M1601190000 
            Caption         =   "&Calculo de Planillas"
            Index           =   0
         End
         Begin VB.Menu M1601190000 
            Caption         =   "&Abono a Cuenta"
            Index           =   1
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Cierre &Mensual"
            Index           =   2
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Cierre &Diario"
            Index           =   3
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Reportes"
         Index           =   31
         Begin VB.Menu M1601200000 
            Caption         =   "&Reportes"
            Index           =   0
         End
         Begin VB.Menu M1601200000 
            Caption         =   "Reportes &Generales"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M1700000000 
      Caption         =   "&Logística"
      Begin VB.Menu M1701000000 
         Caption         =   "Usuarios"
         Index           =   0
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M1701000000 
         Caption         =   "&Bienes"
         Index           =   2
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu M1701000000 
         Caption         =   "&Proveedores"
         Index           =   4
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Requerimiento Regular"
         Index           =   6
         Begin VB.Menu mnuLogRequerimiento 
            Caption         =   "&Solicitud"
         End
         Begin VB.Menu mnuLogTramite 
            Caption         =   "&Trámite"
         End
         Begin VB.Menu mnuLogRegPreRef 
            Caption         =   "Precios Referenciales"
         End
         Begin VB.Menu mnuLogAprobacion 
            Caption         =   "Aprobación o Rechazo"
         End
         Begin VB.Menu mnuLogCtaCnt 
            Caption         =   "Asignación de Ctas.Presupuestrales"
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Requerimiento &Extemporaneo"
         Index           =   7
         Begin VB.Menu mnuLogExtReq 
            Caption         =   "Solicitud"
         End
         Begin VB.Menu mnuLogExtTra 
            Caption         =   "Trámite"
         End
         Begin VB.Menu mnuLogExtPreRef 
            Caption         =   "Precio Referencial"
         End
         Begin VB.Menu mnuLogExtCtaCnt 
            Caption         =   "Asignación de Ctas.Presupuestrales"
         End
         Begin VB.Menu mnuLogExtApro 
            Caption         =   "Aprobacion o Rechazo"
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         Index           =   9
         Begin VB.Menu mnuLogAdquiConsol 
            Caption         =   "Consolidación de los Requerimientos al plan Anual de Adquisiciones "
         End
         Begin VB.Menu mnuLogAdquiConsul 
            Caption         =   "Consultas"
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Proceso de Selección"
         Index           =   11
         Begin VB.Menu mnuLogDefPar 
            Caption         =   "Definición de Parámetros"
         End
         Begin VB.Menu mnuLogSelecIni 
            Caption         =   "Inicio de proceso"
            Begin VB.Menu mnuLogSelecIni1 
               Caption         =   "Resolución o acuerdo de inicio"
               Begin VB.Menu mnuLogSelIni 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelIni 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni2 
               Caption         =   "Conformación del Comite especial u organo encargado"
               Begin VB.Menu mnuLogSelCom 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelCom 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni3 
               Caption         =   "Detalle del proceso"
               Begin VB.Menu mnuLogSelBas 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelBas 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni4 
               Caption         =   "Bases o terminos de referencia"
               Begin VB.Menu mnuLogSelPar 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelPar 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecCon 
            Caption         =   "Convocatoria"
            Begin VB.Menu mnuLogSelEcConP 
               Caption         =   "Publicación"
               Begin VB.Menu mnuLogSelPub 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelPub 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelEcConC 
               Caption         =   "Solicitudes de cotización"
               Begin VB.Menu mnuLogSelCot 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelCot 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecBas 
            Caption         =   "Bases"
            Begin VB.Menu mnuLogSelecBase 
               Caption         =   "Registro de entrega de Bases"
               Begin VB.Menu mnuLogSelecBases 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecBases 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecConBas 
               Caption         =   "Consulta a Bases"
               Begin VB.Menu mnuLogSelecConsultaBase 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecConsultaBase 
                  Caption         =   "Cosulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecAbsolu 
               Caption         =   "Absolución de Consultas"
               Begin VB.Menu mnuLogSelecAbsolucion 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecAbsolucion 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecObserva 
               Caption         =   "Registro de Observaciones"
               Begin VB.Menu mnuLogSelecObsBases 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecObsBases 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecPro 
            Caption         =   "Propuestas"
            Begin VB.Menu mnuLogSelTec 
               Caption         =   "Técnica"
            End
            Begin VB.Menu mnuLogSelecO 
               Caption         =   "Económica"
            End
            Begin VB.Menu mnuLogSelGar 
               Caption         =   "Garantía de seriedad de oferta"
            End
         End
         Begin VB.Menu mnuLogSelecEva 
            Caption         =   "Evaluaciones"
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Proceso de Selección Desierto"
         Index           =   12
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Cancelacion Proceso Seleccion"
         Index           =   13
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Adjudicación del Proceso de Selección"
         Index           =   14
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Consentimiento de Adjudicación"
         Index           =   15
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Consultas"
         Index           =   16
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Contratación"
         Index           =   18
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Mantenimiento Contratación"
            Index           =   1
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Impresión de Contratación"
            Index           =   2
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Provisiones pago de Proveedores"
         Index           =   19
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Almacén"
         Index           =   21
         Begin VB.Menu mnuLogAlmBieSer 
            Caption         =   "Bienes suministros y servicios contratados"
            Begin VB.Menu mnuLogAlmBieSerRecep 
               Caption         =   "Recepción"
            End
            Begin VB.Menu mnuLogAlmBieSerVeri 
               Caption         =   "Verificación de recepción"
               Begin VB.Menu mnuLogAlmBieSerVeriSoli 
                  Caption         =   "Solicitud"
               End
               Begin VB.Menu mnuLogAlmBieSerVeriVeri 
                  Caption         =   "Verificación"
               End
            End
            Begin VB.Menu mnuLogAlmBieSerDistri 
               Caption         =   "Distribución"
            End
            Begin VB.Menu mnuLogAlmBieSerTransfe 
               Caption         =   "Transferencia entre almacenes"
            End
            Begin VB.Menu mnuLogAlmBieSerInven 
               Caption         =   "Inventario"
            End
         End
         Begin VB.Menu mnuLogAlmBieCorrectivo 
            Caption         =   "Bienes para mantenimiento correctivo"
         End
         Begin VB.Menu mnuLogAlmBieDesuso 
            Caption         =   "Bienes en desuso, obsolescencia, malogrados"
         End
         Begin VB.Menu mnuLogAlmBieBaja 
            Caption         =   "Bienes de Baja"
         End
         Begin VB.Menu mnuLogAlmBieDacion 
            Caption         =   "Bienes en dación de pago"
         End
         Begin VB.Menu mnuLogAlmBieEmbargado 
            Caption         =   "Bienes Embargados"
         End
         Begin VB.Menu mnuLogAlmBieAdjudicado 
            Caption         =   "Bienes Adjudicados"
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Depreciación de Bienes"
         Index           =   22
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Remate de Bienes"
         Index           =   23
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Adjudicación de Bienes"
         Index           =   24
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Subasta de Bienes"
         Index           =   25
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Servicios (locación, públicos, privados, móviles, otros)"
         Index           =   27
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Distribución"
            Index           =   1
         End
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Garantías"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub M0101000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmImpresora.Show 1
        Case 1
            If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                End
            End If
    End Select
    
End Sub

Private Sub M0201000000_Click(Index As Integer)
    Select Case Index
    Case 7 'Refinanciacion de Credito
        Call frmCredSolicitud.RefinanciaCredito(Registrar)
    Case 8 ' 'Actualizacion con Metodos de Liquidacion
        frmCredMntMetLiquid.Show 1
    End Select
End Sub

Private Sub M0201010100_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Parametros
            frmCredMantParametros.InicioActualizar
        Case 1 'Consulta de Parametros
            frmCredMantParametros.InicioCosultar
    End Select
End Sub

Private Sub M0201010200_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Lineas de Credito
            frmCredLineaCredito.Registrar
        Case 1 'Mantenimiento de Lineas de Credito
            frmCredLineaCredito.Actualizar
        Case 2 ' Consulta de lineas de Credito
            frmCredLineaCredito.Consultar
    End Select
End Sub

Private Sub M0201010300_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesActualizar
        Case 1 'Consulta de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesConsulta
    End Select
End Sub

Private Sub M0201010400_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Gastos
            frmCredMntGastos.Inicio InicioGastosActualizar
        Case 1
            frmCredMntGastos.Inicio InicioGastosConsultar
    End Select
End Sub

Private Sub M0201020000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Solicitud
            frmCredSolicitud.Inicio Registrar
        Case 1 'Consulta de Solicitud
            frmCredSolicitud.Inicio Consulta
    End Select
End Sub

Private Sub M0201030000_Click(Index As Integer)
Dim oCredRel As New UCredRelacion

    Select Case Index
        Case 0 'Mantenimiento de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
            Set oCredRel = Nothing
        Case 1 'Reasignacion de Cartera en Lote
            frmCredReasigCartera.Show 1
        Case 2 'Consulta de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
            Set oCredRel = Nothing
    End Select
End Sub

Private Sub M0201040000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Garantia
            frmPersGarantias.Inicio RegistroGarantia
        Case 1 'Mantenimiento de Garantia
            frmPersGarantias.Inicio MantenimientoGarantia
        Case 2 'Consulta de Garantia
            frmPersGarantias.Inicio ConsultaGarant
        Case 3 'Gravament
            frmCredGarantCred.Inicio PorMenu
    End Select
End Sub

Private Sub M0201050000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Sugerencia
            frmCredSugerencia.Show 1
        Case 1 'Mantenimiento de Sugerencia
            
        Case 2 'Consulta de Sugerencia
            
    End Select
End Sub

Private Sub M0201060000_Click(Index As Integer)
    Select Case Index
        Case 0 'Aprobacion de Credito
            frmCredAprobacion.Show 1
        Case 1 'Rechazo de Credito
            frmCredRechazo.Rechazar
        Case 2 'Anulacion de Credito
            frmCredRechazo.Retirar
    End Select
End Sub

Private Sub M0201070000_Click(Index As Integer)
    Select Case Index
        Case 0 'Reprogramacion de Credito
            frmCredReprogCred.Show 1
        Case 1 'Reprogramacion en Lote
            frmCredReprogLote.Show 1
    End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub mnuCapBenCons04070100_Click()
    frmCapBeneficiario.Inicia True
End Sub

Private Sub mnuCapBenMnt04070200_Click()
    frmCapBeneficiario.Inicia False
End Sub

Private Sub mnuCapBloqAho04040100_Click()
    frmCapBloqueoDesbloqueo.Inicia gCapAhorros
End Sub

Private Sub mnuCapBloqCTS04040300_Click()
    frmCapBloqueoDesbloqueo.Inicia gCapCTS
End Sub

Private Sub mnuCapBloqPF04040200_Click()
    frmCapBloqueoDesbloqueo.Inicia gCapPlazoFijo
End Sub

Private Sub mnuCapMntAho04030100_Click()
    frmCapMantenimiento.Inicia gCapAhorros
End Sub

Private Sub mnuCapMntCTS04030300_Click()
    frmCapMantenimiento.Inicia gCapCTS
End Sub

Private Sub mnuCapMntPF04030200_Click()
    frmCapMantenimiento.Inicia gCapPlazoFijo
End Sub

Private Sub mnuCapOrdAnu04080300_Click()
    frmCapOrdPagAnulCert.Inicia gAhoOPAnulacion
End Sub

Private Sub mnuCapOrdCer04080200_Click()
    frmCapOrdPagAnulCert.Inicia gAhoOPCertificacion
End Sub

Private Sub mnuCapOrdCon04080400_Click()
    frmCapOrdPagConsulta.Show 1
End Sub

Private Sub mnuCapOrdGen04080100_Click()
    frmCapOrdPagGenEmi.Show 1
End Sub

Private Sub mnuCapParCons04010100_Click()
    frmCapParametros.Inicia True
End Sub

Private Sub mnuCapParMnt04010200_Click()
    frmCapParametros.Inicia False
End Sub

Private Sub mnuCapSerConvCta04090402_Click()
    frmCapServConvCuentas.Inicia
End Sub

Private Sub mnuCapSerConvMnt04090401_Click()
    frmCapConvenioMant.Show 1
End Sub

Private Sub mnuCapSerConvPln04090403_Click()
    frmCapServConvPlanPag.Inicia
End Sub

Private Sub mnuCapSerHidPar04090201_Click()
    frmCapServParametros.Inicia gCapServHidrandina
End Sub

Private Sub mnuCapSerHidRep04090202_Click()
    frmCapServGenReporte.Inicia gCapServHidrandina
End Sub

Private Sub mnuCapSerSedGen04090103_Click()
    frmCapServGeneraDBF.Show 1
End Sub

Private Sub mnuCapSerSedPar04090101_Click()
    frmCapServParametros.Inicia gCapServSedalib
End Sub

Private Sub mnuCapSerSedRep04090102_Click()
    frmCapServGenReporte.Inicia gCapServSedalib
End Sub

Private Sub mnuCapSimPF04050000_Click()
    frmCapSimulacionPF.Show 1
End Sub

Private Sub mnuCapTarBloq04060300_Click()
    frmCapTarjetaBlqDesBlq.Show 1
End Sub

Private Sub mnuCapTarCla04060400_Click()
    frmCapTarjetaCambioClave.Show 1
End Sub

Private Sub mnuCapTarReg04060100_Click()
    frmCapTarjetaRegistro.Inicia
End Sub

Private Sub mnuCapTarRel04060200_Click()
    frmCapTarjetaRelacion.Inicia False
End Sub

Private Sub mnuCapTasaCTSCons04020301_Click()
    frmCapTasaInt.Inicia gCapCTS, True
End Sub

Private Sub mnuCapTasaCTSCons04020302_Click()
    frmCapTasaInt.Inicia gCapCTS, False
End Sub

Private Sub mnuCapTasAhoCons04020101_Click()
    frmCapTasaInt.Inicia gCapAhorros, True
End Sub

Private Sub mnuCapTasAhoMnt04020102_Click()
    frmCapTasaInt.Inicia gCapAhorros, False
End Sub

Private Sub mnuCapTasaPFCons04020201_Click()
    frmCapTasaInt.Inicia gCapPlazoFijo, True
End Sub

Private Sub mnuCapTasaPFCons04020202_Click()
    frmCapTasaInt.Inicia gCapPlazoFijo, False
End Sub

Private Sub mnuCliPosic08010000_Click()
    frmPosicionCli.Show 1
End Sub

Private Sub mnucredAnaMetas02150200_Click()
    frmCredMetasAnalista.Show 1
End Sub

Private Sub mnucredAnaNota02150100_Click()
    frmCredAsigNota.Show 1
End Sub

Private Sub mnucredGastosLote02110100_Click()
    frmCredAsigGastosLote.Show 1
End Sub

Private Sub mnucredGastosPenalidad02110200_Click()
    frmCredExonerarPen.Show 1
End Sub

Private Sub mnucredHistorial02160100_Click()
    frmCredConsulta.Show 1
End Sub

Private Sub mnucredPasarAJud02140000_Click()
    frmCredTransARecup.Show 1
End Sub

Private Sub mnucredPerdMora02100000_Click()
    frmCredPerdonarMora.Show 1
End Sub



Private Sub mnucredReasignarInst02130000_Click()
    frmCredReasigInst.Show 1
End Sub



Private Sub mnucredRepDupDocum02170100_Click()
    frmCredDupDoc.Show 1
End Sub


Private Sub mnucredSimCalCuoLib02090300_Click()
Dim MatCalend As Variant
Dim Matriz(0) As String
    MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
End Sub

Private Sub mnucredSimCalDesPar02090200_Click()
    frmCredCalendPagos.Simulacion DesembolsoParcial
End Sub

Private Sub mnucredSimCalPag02090100_Click()
    frmCredCalendPagos.Simulacion DesembolsoTotal
End Sub



Private Sub mnuHerramPerif09020000_Click()
    frmSetupCOM.Show 1
End Sub

Private Sub mnuHerramSpooler09010000_Click()
    frmSpooler.Show 1
End Sub

Private Sub mnuIFinanCons07030200_Click()
    frmMntInstFinanc.InicioConsulta
End Sub


Private Sub mnuOpCajero06900000_Click()
    frmCajeroOperaciones.Show 1
End Sub

Private Sub mnuOpCajeroCMACLLam06920000_Click()
    frmCajeroOpeCMAC.Inicia False
End Sub

Private Sub mnuOpCajeroCMACRes06910000_Click()
    frmCajeroOpeCMAC.Inicia
End Sub

Private Sub mnuOpCajeroExtCapta06930400_Click()
    frmCapExtornos.Show 1
End Sub

Private Sub mnuOpeDesemAbo06010000_Click()
    frmCredDesembAbonoCta.DesembolsoCargoCuenta
End Sub

Private Sub mnupersonaCons07010300_Click()
    frmPersona.Consultar
End Sub

Private Sub mnupersonamant07010200_Click()
    frmPersona.Mantenimeinto
End Sub

Private Sub mnupersonareg07010100_Click()
    frmPersona.Registrar
End Sub

Private Sub mnuPigAdjudica03040000_Click()
    frmColPAdjudicaLotes.Show 1
End Sub

Private Sub mnuPigAnulacion03010300_Click()
    frmColPAnularPrestamoPig.Show
End Sub

Private Sub mnuPigBloqueo03010400_Click()
    frmColPBloqueo.Show 1
End Sub

Private Sub mnuPigContraRegis03010100_Click()
    frmColPRegContrato.Show 1
End Sub

Private Sub mnuPigMntgDescrip03010200_Click()
    frmColPMantPrestamoPig.Show
End Sub

Private Sub mnuPigRemPrepRem03030100_Click()
     frmColPRematePrepara.Show 1
End Sub

Private Sub mnuPigRemRem03030200_Click()
    frmColPRemateProceso.Show 1
End Sub

Private Sub mnuPigRescate03020000_Click()
    frmColPRescateJoyas.Show 1
End Sub

Private Sub mnuPigSubPrep03050100_Click()
    frmColPSubastaPrepara.Show 1
End Sub

Private Sub mnuPigSubSubasta03050200_Click()
    frmColPSubastaProceso.Show 1
End Sub

Private Sub mnuSegurPerm10010000_Click()
    frmMantPermisos.Show 1
End Sub

Private Sub mnutFinanMant07030100_Click()
    frmMntInstFinanc.InicioActualizar
End Sub

Private Sub mnuLogSerServicios_Click(Index As Integer)

End Sub

Private Sub Tiempo_Timer()
    SBBarra.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & Space(3) & Format(Time, "hh:mm AMPM")
End Sub
