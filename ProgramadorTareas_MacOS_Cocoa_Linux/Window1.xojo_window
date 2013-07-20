#tag Window
Begin Window Window1
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   406
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   1969461247
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   1
   Resizeable      =   False
   Title           =   "Programador de Tareas"
   Visible         =   True
   Width           =   767
   Begin GroupBox GroupBox4
      AutoDeactivate  =   True
      Bold            =   False
      Caption         =   ""
      Enabled         =   True
      Height          =   74
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   347
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   12
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   258
      Underline       =   False
      Visible         =   True
      Width           =   108
      Begin PushButton VerLog
         AutoDeactivate  =   True
         Bold            =   False
         ButtonStyle     =   "0"
         Cancel          =   False
         Caption         =   "Ver Log"
         Default         =   False
         Enabled         =   True
         Height          =   22
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox4"
         Italic          =   False
         Left            =   358
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0.0
         TextUnit        =   0
         Top             =   299
         Underline       =   False
         Visible         =   True
         Width           =   87
      End
      Begin PushButton BorrarTarea
         AutoDeactivate  =   True
         Bold            =   False
         ButtonStyle     =   "0"
         Cancel          =   False
         Caption         =   "Borrar Tarea"
         Default         =   False
         Enabled         =   True
         Height          =   22
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox4"
         Italic          =   False
         Left            =   358
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   0
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0.0
         TextUnit        =   0
         Top             =   274
         Underline       =   False
         Visible         =   True
         Width           =   87
      End
   End
   Begin GroupBox GroupBox1
      AutoDeactivate  =   True
      Bold            =   False
      Caption         =   "Nueva Tarea"
      Enabled         =   True
      Height          =   58
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   27
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   258
      Underline       =   False
      Visible         =   True
      Width           =   308
      Begin PushButton InsertarTarea
         AutoDeactivate  =   True
         Bold            =   False
         ButtonStyle     =   "0"
         Cancel          =   False
         Caption         =   "Insertar Tarea"
         Default         =   False
         Enabled         =   True
         Height          =   22
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox1"
         Italic          =   False
         Left            =   244
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   0
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0.0
         TextUnit        =   0
         Top             =   284
         Underline       =   False
         Visible         =   True
         Width           =   80
      End
      Begin TextField NuevaTarea
         AcceptTabs      =   False
         Alignment       =   0
         AutoDeactivate  =   True
         AutomaticallyCheckSpelling=   False
         BackColor       =   &cFFFFFF00
         Bold            =   False
         Border          =   True
         CueText         =   ""
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Format          =   ""
         Height          =   22
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox1"
         Italic          =   False
         Left            =   39
         LimitText       =   0
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Mask            =   ""
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextColor       =   &c00000000
         TextFont        =   "System"
         TextSize        =   0.0
         TextUnit        =   0
         Top             =   284
         Underline       =   False
         UseFocusRing    =   True
         Visible         =   True
         Width           =   193
      End
   End
   Begin DataControl DataControl1
      AutoDeactivate  =   True
      Bold            =   False
      Caption         =   "Sin título"
      Commit          =   True
      Database        =   "902307839"
      Enabled         =   True
      Height          =   25
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   537
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      ReadOnly        =   False
      Scope           =   0
      SQLQuery        =   "select * from tareas;"
      TabIndex        =   3
      TableName       =   "Tareas"
      TabPanelIndex   =   0
      TabStop         =   True
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   13
      Underline       =   False
      Visible         =   False
      Width           =   200
   End
   Begin GroupBox GroupBox2
      AutoDeactivate  =   True
      Bold            =   False
      Caption         =   "Log Ultima Ejecucion"
      Enabled         =   True
      Height          =   136
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   467
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   258
      Underline       =   False
      Visible         =   True
      Width           =   280
   End
   Begin Listbox ListaTareas
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   7
      ColumnsResizable=   True
      ColumnWidths    =   "50,150,80,60,60,180"
      DataField       =   "NombreTarea"
      DataSource      =   "DataControl1"
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   False
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   True
      HeadingIndex    =   -1
      Height          =   197
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   0
      ScrollbarHorizontal=   True
      ScrollBarVertical=   True
      SelectionType   =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   53
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   727
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin Timer Timer1
      Height          =   32
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   560
      LockedInPosition=   False
      Mode            =   2
      Period          =   60000
      Scope           =   0
      TabPanelIndex   =   "0"
      Top             =   40
      Width           =   32
   End
   Begin TextField Tiempo
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   8
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   13
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   False
      Width           =   63
   End
   Begin TextField Tiempo2
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   138
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   13
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   False
      Width           =   59
   End
   Begin TextArea TextArea1
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   True
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   112
      HelpTag         =   ""
      HideSelection   =   True
      Index           =   -2147483648
      Italic          =   False
      Left            =   477
      LimitText       =   0
      LineHeight      =   0.0
      LineSpacing     =   1.0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Multiline       =   True
      ReadOnly        =   True
      Scope           =   0
      ScrollbarHorizontal=   True
      ScrollbarVertical=   True
      Styled          =   True
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   274
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   260
   End
   Begin TextField Tiempo1
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   73
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   13
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   False
      Width           =   60
   End
   Begin PushButton EjecutarTimer
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Ejecutar Timer Manual"
      Default         =   False
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   347
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   240
      Underline       =   False
      Visible         =   False
      Width           =   129
   End
   Begin Label StaticText1
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   34
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   10
      TabPanelIndex   =   0
      Text            =   "Programador de Tareas"
      TextAlign       =   1
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   24.0
      TextUnit        =   0
      Top             =   7
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   727
   End
   Begin GroupBox GroupBox3
      AutoDeactivate  =   True
      Bold            =   False
      Caption         =   "Ayuda"
      Enabled         =   True
      Height          =   61
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   27
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   11
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   333
      Underline       =   False
      Visible         =   True
      Width           =   428
      Begin Label TextoAyuda
         AutoDeactivate  =   True
         Bold            =   False
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Height          =   37
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox3"
         Italic          =   False
         Left            =   39
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   True
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   0
         Text            =   ""
         TextAlign       =   0
         TextColor       =   &c00000000
         TextFont        =   "System"
         TextSize        =   12.0
         TextUnit        =   0
         Top             =   349
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   409
      End
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  
		  base.Append(New MenuItem ("Acerca de"))
		  
		  
		End Function
	#tag EndEvent

	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  
		  Acercade.Show
		  
		End Function
	#tag EndEvent

	#tag Event
		Sub Open()
		  
		  //Dim db As New  SQLiteDatabase
		  Dim db As New  REALSQLDatabase
		  Dim dbFile As FolderItem = GetFolderItem("ProgramadorTareas.sqlite")
		  Dim ControlDatos as new DataControl
		  dim TextoSQL as string
		  
		  If dbFile <> Nil And dbFile.Exists Then
		    db.DatabaseFile = dbFile
		    
		    If db.Connect Then
		      
		      db.DatabaseFile=dbFile
		      ControlDatos.Database=db
		      Window1.Datacontrol1.Database=ControlDatos.database
		      Window1.Datacontrol1.TableName="Tareas"
		      Window1.Datacontrol1.SQLQuery="select * from Tareas;"
		      Window1.Datacontrol1.RunQuery
		      
		    end if
		    
		  else
		    
		    'Crea Base de Datos en directorio actual
		    db.DatabaseFile = dbFile
		    
		    dbFile = New FolderItem("ProgramadorTareas.sqlite")
		    db.DatabaseFile = dbFile
		    If not db.CreateDatabaseFile Then
		      MsgBox("Database not created. Error: " + db.ErrorMessage)
		      exit 
		    End If
		    
		    TextoSQL="CREATE TABLE [Tareas] ([IdTarea] INTEGER  PRIMARY KEY AUTOINCREMENT NOT NULL,[NombreTarea] VarChar(50)  NOT NULL,[DiaSemana] VarChar  NULL,[DiaMes] VarChar(2)  NULL,[Hora] VARCHAR(5)  NULL,[ComandoaEjecutar] VARCHAR(255)  NULL,[UltimaEjecucionTarea] VARCHAR(20)  NULL);"
		    db.SQLExecute TextoSQL
		    db.Commit
		    db.DatabaseFile = dbFile
		    ControlDatos.Database=db
		    Window1.Datacontrol1.Database=ControlDatos.database
		    Window1.Datacontrol1.TableName="Tareas"
		    Window1.Datacontrol1.SQLQuery="select * from Tareas;"
		    Window1.Datacontrol1.RunQuery
		    
		  end if
		  
		  goto llenarlista
		  
		  'Esto no es necesario para que funcione correctamente, insertaría un primer registro en la tabla Tareas
		  'if Window1.Datacontrol1.RecordSet.RecordCount=0 then 'Hay que insertar al menos un registro
		  'TextoSQL="insert into Tareas values(NULL,'Tarea Ejemplo','','','','','','');"
		  'db.SQLExecute TextoSQL
		  'db.Commit
		  'Window1.Datacontrol1.RunQuery
		  'end if
		  
		  llenarlista:
		  
		  OperarConLista(datacontrol1, ListaTareas,"Llenar","",0,0,"")
		  'ListaTareas.CellType(2,0)=ListaTareas.TypePopupMenu
		  
		  ListaTareas.ColumnType(2)=1 'Solo lectura en DiaSemana
		  
		  Window1.Show
		  
		  
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub EjecutarTareas()
		  
		  Dim NumFila as Integer
		  
		  dim hora as integer, minuto as integer, segundo as integer
		  'dim hora2 as integer, minuto2 as integer
		  dim minutoprocesado as integer
		  dim fecha as new date
		  dim DiaSemana as integer, DiaSemana2 as String
		  Dim DiaMes as integer, DiaMes2 as integer
		  
		  hora=fecha.Hour
		  minuto=fecha.minute
		  segundo=fecha.second
		  DiaSemana=fecha.DayOfWeek
		  DiaMes=fecha.Day
		  
		  'msgbox "Dia Semana: " + str(DiaSemana)
		  
		  Tiempo.Text=str(hora)+":"+str(minuto)+":"+str(segundo)
		  
		  Static HoraCorta as string
		  //HoraCorta=str(fecha.hour,"00") + ":" + str(fecha.minute,"00")
		  HoraCorta=format(fecha.hour,"00") + ":" + format(fecha.minute,"00")
		  
		  'minutoprocesado=99 'Para que nunca se de ´´ solo haria entonces uno
		  'try
		  'if datacontrol1.recordset.EOF then
		  'exit
		  'end if
		  //if datacontrol1.recordset=Nil then 'No hay registros, o no hay base de datos
		  if datacontrol1.recordset.recordcount=0 then 'No hay registros, o no hay base de datos
		    exit
		  end if
		  datacontrol1.recordset.movefirst 'Para ejecutar esta sentencia debe de haber una base de datos
		  
		  NumFila=0  'Para posicionarnos luego en la tarea correspondiente en la lista
		  
		  do  until datacontrol1.recordset.eof
		    
		    'hora2=datacontrol1.recordset.field("Hora").Value
		    'minuto2=datacontrol1.recordset.field("Minuto").Value
		    Static HoraCorta2 as string
		    if datacontrol1.recordset.field("Hora").Value="" or datacontrol1.recordset.field("Hora").Value.IsNull then
		      HoraCorta2=""
		    else
		      if datacontrol1.recordset.field("Hora").Value="" or datacontrol1.recordset.field("Hora").Value.IsNull then
		        HoraCorta2=""
		      else
		        if len(datacontrol1.recordset.field("Hora").Value)<5 then
		          HoraCorta2="0"+datacontrol1.recordset.field("Hora").Value
		        else
		          HoraCorta2=datacontrol1.recordset.field("Hora").Value
		        end if
		      end if
		    end if
		    
		    DiaSemana2=datacontrol1.recordset.field("DiaSemana").Value
		    
		    tiempo1.text=HoraCorta 'str(hora)+":"+str(minuto)
		    tiempo2.text=HoraCorta2 'str(hora2) + ":"+str(minuto2)
		    
		    static FechaHora as string
		    FechaHora=fecha.ShortDate + "  " +fecha.ShortTime
		    
		    static ProcesaDia as boolean =false
		    
		    select case DiaSemana2
		    case "Todos"
		      ProcesaDia=true
		    case "Lunes"
		      if DiaSemana=2 then
		        ProcesaDia=true
		      end if
		    case "Martes"
		      if Diasemana=3 then
		        ProcesaDia=true
		      end if
		    case "Miercoles"
		      if DiaSemana=4 then
		        ProcesaDia=true
		      end if
		    case "Jueves"
		      if DiaSemana=5 then
		        ProcesaDia=true
		      end if
		    case "Viernes"
		      if DiaSemana=6 then
		        ProcesaDia=true
		      end if
		    case "Sabado"
		      if DiaSemana=7 then
		        ProcesaDia=true
		      end if
		    case "Domingo"
		      if DiaSemana=1 then
		        ProcesaDia=true
		      end if
		    case "L-V"
		      if DiaSemana=2 or DiaSemana=3 or DiaSemana=4 or DiaSemana=5 or DiaSemana=6 then
		        ProcesaDia=True
		      end if
		    end select
		    
		    'DIA MES
		    if datacontrol1.recordset.field("DiaMes").Value="" or datacontrol1.recordset.field("DiaMes").Value.IsNull then
		      DiaMes2=0
		    else
		      DiaMes2=datacontrol1.recordset.field("DiaMes").Value
		    end if
		    
		    Static ProcesaDiaMes as boolean=false
		    
		    if DiaMes2=0 then
		      ProcesaDiaMes=true
		    else
		      if DiaMes=DiaMes2 then
		        ProcesaDiaMes=true
		      end if
		    end if
		    
		    if ProcesaDia=true and ProcesaDiaMes=true and HoraCorta=HoraCorta2 and not (minutoprocesado=minuto) then
		      
		      static textotarea as String, TextoLog as String, TextoComando as String
		      
		      textotarea= datacontrol1.recordset.field("NombreTarea")
		      TextoLog=FechaHora + Chr(13) + "TAREA: " + textotarea + Chr(13)
		      EscribirLog("ProgramadorTareas.log","--------------------------------------------" + chr(13))
		      EscribirLog("ProgramadorTareas.log",TextoLog+chr(13))
		      
		      'msgbox ("Procesa programación. Nombre Tarea:" + textotarea)
		      
		      'minutoprocesado=minuto
		      'Ejecuta el Shell
		      Dim sh As New Shell
		      
		      TextoComando=datacontrol1.recordset.field("ComandoaEjecutar")
		      TextoLog="COMANDO A EJECUTAR: " + Chr(13) +TextoComando + Chr(13)
		      EscribirLog("ProgramadorTareas.log",TextoLog+chr(13))
		      
		      sh.Execute(datacontrol1.recordset.field("ComandoaEjecutar"))
		      TextArea1.Text = sh.Result
		      TextoLog=FechaHora +  Chr(13) + "RESULTADO: " + Chr(13) +TextArea1.Text + Chr(13)
		      EscribirLog("ProgramadorTareas.log",TextoLog+chr(13))
		      
		      Static SentenciaSQL as string
		      
		      SentenciaSQL="UPDATE Tareas SET UltimaEjecucionTarea='" + FechaHora + "' WHERE " + datacontrol1.recordset.Field(ListaTareas.Heading(0)).Name + " = " +  ListaTareas.Cell(ListaTareas.Listindex,0) + "; COMMIT;"
		      'msgbox SentenciaSQL
		      //ControldeDatos.recordset.Update
		      datacontrol1.Database.SQLExecute SentenciaSQL
		      If datacontrol1.Database.Error Then
		        MsgBox("DB Error: " + datacontrol1.Database.ErrorMessage)
		      End If
		      
		      Datacontrol1.RunQuery
		      
		      'datacontrol1.recordset.edit
		      'datacontrol1.recordset.Field("UltimaEjecucionTarea").Value=FechaHora
		      'datacontrol1.recordset.update
		      
		      'ListaTareas.Cell(NumFila,8)=FechaHora
		      
		      //datacontrol1.RunQuery  'Para que actualice el datacontrol principal
		      OperarConLista(Datacontrol1,ListaTareas,"Llenar","",0,0,"")
		      
		    end if
		    
		    NumFila=NumFila+1
		    datacontrol1.recordset.movenext
		  loop
		  
		End Sub
	#tag EndMethod


#tag EndWindowCode

#tag Events VerLog
	#tag Event
		Sub Action()
		  
		  
		  VentanaLog.Show
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events BorrarTarea
	#tag Event
		Sub Action()
		  
		  if ListaTareas.Listindex<>-1 then
		    OperarConLista(Datacontrol1, ListaTareas, "BorrarFila","Tareas",ListaTareas.Listindex,0,"")
		  else
		    msgbox "Debe de seleccionar una tarea para poder borrarla"
		  end if
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events InsertarTarea
	#tag Event
		Sub Action()
		  
		  
		  if trim(NuevaTarea.Text)<>"" then
		    OperarConLista(Datacontrol1, ListaTareas, "InsertarFila",NuevaTarea.Text,0,0,"")
		  else
		    msgbox "Debe de insertar un nombre de tarea para poder crearla"
		  end if
		  
		  NuevaTarea.Text=""
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ListaTareas
	#tag Event
		Sub CellLostFocus(row as Integer, column as Integer)
		  
		  
		  if ListaTareas.Heading(Column)="Hora" then
		    static Campo as string
		    Campo=ListaTareas.Cell(row,column)
		    ListaTareas.Cell(row,column)=FormatearCampo("Texto","HoraCorta",Campo)
		  end if
		  
		  OperarConLista(Datacontrol1,ListaTareas,"ActualizarCelda","Tareas", row,column,ListaTareas.Cell(row,column))
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  
		  'base.Append(New MenuItem ("Lunes"))
		  'base.Append(New MenuItem ("Martes"))
		  'base.Append(New MenuItem ("Miercoles"))
		  'base.Append(New MenuItem ("Jueves"))
		  'base.Append(New MenuItem ("Viernes"))
		  'base.Append(New MenuItem ("Sabado"))
		  'base.Append(New MenuItem ("Domingo"))
		  'base.Append(New MenuItem ("L-V"))
		  'base.Append(New MenuItem ("Todos"))
		  
		End Function
	#tag EndEvent
	#tag Event
		Sub CellAction(row As Integer, column As Integer)
		  
		  
		  '
		  'If me.CellType(row, column) = me.TypePopupMenu  or me.ColumnType(column) = me.TypePopupMenu then
		  '
		  'Dim m as New MenuItem
		  '
		  'm.append( New MenuItem("More Data"))
		  'm.append( New MenuItem("Less Data"))
		  'm.append( New MenuItem(MenuItem.TextSeparator))
		  '
		  'm.append( New MenuItem("More or less"))
		  'm.append( New MenuItem("No Data"))
		  'm.append( New MenuItem("Infinite Data"))
		  '
		  'm = m.PopUp(PopupX, PopupY)
		  'End If
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(x As Integer, y As Integer) As Boolean
		  
		  if ListaTareas.ListIndex=-1 then  'Si no tiene ninguna fila seleccionada, no muestra menu
		    exit
		  end if
		  
		  Dim row As Integer = Me.RowFromXY( x, y )
		  Dim column As Integer = Me.ColumnFromXY ( x, y )
		  Dim Choices() As String
		  
		  Select Case Column
		    
		    'case 0
		    'choices = Array("First","Last")
		    '
		    'Case 1
		    'Choices = Array("Sunday of","Monday of","Tuesday of","Wednesday of","Thursday of","Friday of","Saturday of")
		    
		  Case 2
		    Choices = Array("Lunes","Martes","Miercoles","Jueves","Viernes","Sabado","Domingo","L-V","Todos")
		    
		  end
		  
		  Dim menu As New MenuItem
		  for each choice As String in choices
		    Dim item As New MenuItem( choice )
		    menu.Append item
		    if choice = me.Cell( row, column ) then item.Checked = true
		  next
		  
		  Dim choice As MenuItem = menu.PopUp
		  if choice is nil then return False
		  
		  
		  me.Cell( row, column ) = choice.Text
		  
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  '
		  'if ListaTareas.Listindex=-1 then
		  'msgbox "Debe seleccionar una tarea para asignar día de la semana"
		  'else
		  'dim nombreitem as string
		  'nombreitem=hitItem.Text
		  'ListaTareas.cell(ListaTareas.Listindex,2)=nombreitem
		  'OperarConLista(Datacontrol1,ListaTareas,"ActualizarCelda","",0,0,"")
		  'end if
		  
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  
		  
		  
		End Function
	#tag EndEvent
	#tag Event
		Sub Change()
		  '
		  '
		  'If Keyboard.AsyncKeyDown(124) then
		  'msgbox "Flecha derecha"
		  '//do something with the right arrow key...
		  'static Fila as integer
		  'static Columna as string
		  'Fila=ListaTareas.ListIndex
		  'columna=str(ListaTareas.ActiveCell)
		  'MsgBox "Fila: " + str(fila) + " Columna: " + Columna
		  ''ListaTareas.EditCell(Fila,Columna+1)
		  '
		  'end if
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  
		  if ListaTareas.Listindex=-1 then
		    TextoAyuda.Caption="Debe insertar una nueva tarea"
		  else
		    select case ListaTareas.ColumnFromXY(X,Y)
		    case 0
		      TextoAyuda.Caption="Id Tarea: Autonumérico de sólo lectura"
		    case 1
		      TextoAyuda.Caption="Nombre Tarea: Inserte un nombre para la tarea"
		    case 2
		      TextoAyuda.Caption="Dia Semana: Dia que quiere ejecutar la tarea. L-V de Lunes a Viernes"
		    case 3
		      TextoAyuda.Caption="Dia del mes: Opcional, se puede utilizar en combinación con día de la semana"
		    case 4
		      TextoAyuda.Caption="Hora: Hora de ejecución de la tarea"
		    case 5
		      TextoAyuda.Caption="Comando a ejecutar: Llamada a fichero ejecutable o fichero batch (.bat o fichero por lotes). Ej.: C:\procesar.bat"
		    case 6
		      TextoAyuda.Caption="Última ejecución tarea: Cuando se ejecuta el programador para una tarea, guarda la fecha y hora"
		    case else
		      TextoAyuda.Caption=""
		    end select
		    
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Timer1
	#tag Event
		Sub Action()
		  'dim db as database
		  'dim tabla as DatabaseQuery
		  
		  'db=SQLiteDatabase("ProgramadorTareas")
		  'tabla= SQLiteDatabase ("ProgramadorTareas")
		  'select * from tareas")
		  
		  'tabla.movefirst
		  
		  'tabla.Database=datacontrol1.database.DatabaseName
		  'tabla.database=ProgramadorTareas
		  
		  'tabla.SQLQuery="Select IdTarea from Tareas where IdTarea=1;"
		  
		  'dim campo as DatabaseField
		  
		  'msgbox campo.Value
		  
		  EtiquetaTimer:
		  
		  EjecutarTareas()
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TextArea1
	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  
		  TextoAyuda.Caption="Log Última Ejecución: Cuando se ejecuta una tarea, en esta ventana aparece el último log de ejecución"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events EjecutarTimer
	#tag Event
		Sub Action()
		  'goto EtiquetaTimer
		  EjecutarTareas() 'Ejecuta el código del Timer
		  
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Appearance"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Appearance"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Metal Window"
			"10 - Drawer Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Position"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Appearance"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Position"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Appearance"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Position"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
#tag EndViewBehavior
