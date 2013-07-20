#tag Module
Protected Module Funciones
	#tag Method, Flags = &h0
		Sub EscribirLog(NombreFicheroLog As String, TextoLog as String)
		  
		  //create a binary file with the type of text (defined in the file types dialog)
		  
		  'Dim FicheroBinario as BinaryStream
		  Dim Fichero as FolderItem
		  
		  Dim f as FolderItem
		  Dim stream as BinaryStream
		  Dim fileStream As TextOutputStream
		  
		  'goto Escribir
		  
		  Fichero = GetFolderItem(NombreFicheroLog)
		  'msgbox "Nombre Fichero Log: " + NombreFicheroLog
		  if not Fichero.exists then 'Si no existe el fichero
		    'FicheroBinario=BinaryStream.Create(Fichero,true)
		    fileStream = TextOutputStream.Create(Fichero)
		  else
		    fileStream = TextOutputStream.Append(Fichero)
		    'FicheroBinario= BinaryStream.Open(Fichero, true)
		  end if
		  
		  fileStream.Write(TextoLog)
		  fileStream.Close
		  
		  goto Fin
		  
		  'fileStream.WriteLine(NameField.Text)
		  
		  'Msgbox "Texto Log: " + TextoLog
		  'FicheroBinario.WriteLine(TextoLog)
		  'FicheroBinario.Close
		  
		  
		  Escribir:
		  
		  f = GetSaveFolderItem("*.txt",NombreFicheroLog)
		  If f <> Nil Then
		    stream = BinaryStream.Create(f, True)
		  else
		    stream=BinaryStream.Open(f,True)
		  End If
		  stream.Write(TextoLog)
		  stream.Close
		  
		  Fin:
		  
		  '------------------------------------------------------------------
		  
		  'Crear y sobreescribir
		  'FicheroBinario=BinaryStream.Create(Fichero,true)
		  
		  
		  'else
		  'Abrir de solo lectura
		  'FicheroBinario=BinaryStream.Open(Fichero,false)
		  'end if
		  
		  'FicheroBinario = BinaryStream.Create(Fichero, false)
		  'Fichero= BinaryStream.Open(""+NombreFicheroLog+"", true)
		  'If Not Fichero.   .exists then
		  'Fichero = BinaryStream.create (NombreFicheroLog,true)
		  'end if
		  
		  //check to see if it was created
		  'If Fichero <> Nil Then
		  //write the contents of the editField
		  
		  'Esto es un ejemplo para leer de un Fichero Binario
		  'OutputArea.Text = bs.Read(bs.Length)
		  
		  //close the binaryStream
		  'End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FormatearCampo(TipoCampo as String, TipoFormato as String, Campo as Variant) As Variant
		  
		  
		  'dim a1 as integer= val(mid(tHora.text,1,2))
		  'dim a2 as integer=val(mid(tHora.text,4,2))
		  'dim b as string=Textfield5.text
		  'tHora.Text= str(format(a1, "00")) + ":" + str(format(a2, "00"))
		  
		  Dim Hora as string, Minuto as String
		  Campo=str(Campo)  'Convierte primero a String
		  
		  if TipoCampo="Texto" and TipoFormato="HoraCorta" then
		    if Campo.IsNull or trim(Campo)="" then
		      'Campo="00:00"
		    else
		      static PosicionDosPuntos as integer
		      PosicionDosPuntos=Instr(Campo,":")
		      
		      if PosicionDosPuntos=0 then
		        
		        Hora=trim(Campo)
		        
		        if len(Hora)=1 then
		          Hora="0"+trim(Campo)
		        else 
		          Hora= trim(Campo)
		        end if
		        'Hora=str(Campo,"00")
		        Minuto="00"
		        
		      else
		        static LongitudTexto as integer
		        static CaracteresIzqDosPuntos as integer, TextoDerechaPuntos as string
		        'static CaracteresDerDosPuntos as integer
		        
		        LongitudTexto=Len(Campo)
		        TextoDerechaPuntos=Right(Campo,len(Campo)-PosicionDosPuntos)
		        
		        CaracteresIzqDosPuntos=LongitudTexto-(Len(TextoDerechaPuntos)+1)
		        
		        Hora=mid(Campo,1,CaracteresIzqDosPuntos)
		        Minuto=mid(Campo,CaracteresIzqDosPuntos+2,LongitudTexto-CaracteresIzqDosPuntos-1)
		        if len(Hora)=1 then
		          Hora="0"+Hora
		        end if
		        if len(Minuto)=1 then
		          Minuto="0"+Minuto
		        end if
		        
		        'Campo=str(Hora,"00")+":" + str(Minuto,"00")
		        
		      end if
		      
		      Campo=Hora+":"+Minuto
		      
		    end if
		    
		    '-----------------
		    'Validar Campo
		    '-----------------
		    if val(Hora)>23 or val(Hora)<0 then
		      msgbox "Hora Incorrecta"
		      Campo=""'    "00:"+str(Minuto,"00")
		    end if
		    if val(Minuto)>59 or val(Minuto)<0 then
		      msgbox "Minuto incorrecto"
		      Campo=""   'str(Hora,"00")+":00"
		    end if
		    
		  end if
		  
		  Return Campo
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ObtenerParametro(Parametros as string,NumParametro as integer)
		  
		  'static PosicionComa as integer
		  'static InicioParametro=
		  'static SalirDo=true
		  
		  'do until SalirDo = true
		  
		  'PosicionComa=instr(Parametros,",")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OperarConLista(Datacontrol1 as Datacontrol, Lista1 as Listbox, Operacion as String, Argumentos as String)
		  
		  static i, NumCampos as Integer
		  static NombreCampo as String
		  static ListaCampos as String
		  Static Campos(100) as String //declare array without specifying size
		  
		  Dim ControldeDatos as Datacontrol=Datacontrol1
		  
		  select case Operacion
		  case "Llenar"
		    
		    Lista1.DeleteAllRows
		    
		    if datacontrol1.recordset=Nil then 'No hay registros
		      exit
		    end if
		    datacontrol1.recordset.movefirst
		    
		    NumCampos=Datacontrol1.recordset.FieldCount
		    
		    Lista1.ColumnCount=NumCampos 'Ante de insertar valores
		    
		    'Campos=Array(NombreCampo)
		    
		    Lista1.HasHeading=True
		    
		    for i= 1 to NumCampos
		      NombreCampo=Datacontrol1.recordset.IdxField(i).Name
		      ListaCampos=ListaCampos+chr(13)+NombreCampo
		      'Lista1.Column(i).Name=Datacontrol1.recordset.IdxField(i).Name 'No vale
		      'Lista1.InitialValue(i)=Datacontrol1.recordset.IdxField(i).Name 'No vale
		      'Lista1.AddRow (NombreCampo) 'Esto es para añadir registros/filas
		      'Lista1.InsertRow(i,NombreCampo)
		      Lista1.Heading(i-1)=NombreCampo
		      
		      if i<>1 then 'No hace editable la columna Indice
		        Lista1.ColumnType(i-1)=3 'Hace editable la columna
		      end if
		      
		    next i
		    
		    'Lista1.InitialValue(i)=Datacontrol1.recordset.IdxField(i).Name 'No
		    'Lista1.InitialValue=ListaCampos
		    
		    'static ElementoLista as ListColumn 'Esto puede servir para modificar el ancho de cada columna
		    
		    Do Until datacontrol1.recordset.eof
		      'static ContenidoRow as String = ""
		      'static ContenidoCampo as String=""
		      Static NombreCampo2 as String
		      
		      for i=1 to NumCampos
		        NombreCampo2=str(Datacontrol1.recordset.IdxField(i).Name)
		        'Campos=Array(DataControl1.recordset.field(NombreCampo2) //create 3 elements and assign values
		        Static ContenidoCampo as Variant
		        ContenidoCampo=str(DataControl1.recordset.field(NombreCampo2).Value)
		        Campos(i-1)=ContenidoCampo
		        
		        'ContenidoCampo=DataControl1.recordset.field(Datacontrol1.recordset.IdxField(i).Name)
		        'ContenidoRow=ContenidoRow+","+ ContenidoCampo
		        
		      next i
		      Lista1.AddRow(Campos)
		      'Lista1.AddRow(ContenidoRow)
		      'ListaTareas.AddRow DataControl1.recordset.field("NombreTarea")
		      'DataControl1.Recordset.MoveNext
		      datacontrol1.recordset.movenext
		    Loop
		    '----------------------------------------------------------------
		  case "ActualizarCelda"
		    
		    if Lista1.ListIndex=-1 then 'El usuario no está actualizando la lista, es un proceso, por ejemplo "LlenaLista"
		      exit
		    end if
		    
		    NumCampos=Datacontrol1.recordset.FieldCount
		    
		    'Dim Tabla as recordset=Datacontrol1.recordset
		    'Dim BD as database = Datacontrol1.Database
		    
		    'Tabla.movefirst
		    ControldeDatos.recordset.movefirst
		    ControldeDatos.MoveTo(Lista1.Listindex+1)
		    'msgbox Datacontrol1.recordset.field("NombreTarea").Value
		    'static SentenciaSQL as String
		    'static NombreTabla as String
		    'NombreTabla=Datacontrol1(#kTable)
		    'SentenciaSQL="UPDATE 
		    ControldeDatos.recordset.Edit
		    for i=1 to NumCampos
		      ControldeDatos.recordset.Field(Lista1.Heading(i-1)).Value=Lista1.Cell(Lista1.Listindex,i-1)
		    next i
		    ControldeDatos.recordset.update
		    Datacontrol1.RunQuery
		    
		  case "InsertarFila"
		    
		    static NombreTarea as String =Argumentos
		    
		    Dim row As New DatabaseRecord
		    // ID will be updated automatically
		    row.Column("NombreTarea") = NombreTarea
		    ControldeDatos.database.InsertRecord("Tareas", row)
		    
		    If ControldeDatos.database.Error Then
		      MsgBox("DB Error: " + ControldeDatos.database.ErrorMessage)
		    else
		      'Rellena el nuevo registro en la lista
		      ControldeDatos.RunQuery
		      ControldeDatos.recordset.movelast
		      Lista1.AddRow (ControldeDatos.recordset.Field("IdTarea"),ControldeDatos.recordset.Field("NombreTarea"))
		    End If
		    Datacontrol1.RunQuery
		    
		  case "BorrarFila"
		    
		    static IndiceLista as integer
		    IndiceLista=val(Argumentos)
		    
		    static Confirmacion as integer
		    'static TextoaVisualizar as string
		    
		    
		    'TextoaVisualizar =  Lista1.Cell(IndiceLista,2)
		    
		    Confirmacion=Msgbox("Desea borrar la tarea seleccionada?",1+48+256)
		    if Confirmacion <>1 then
		      exit
		    end if
		    
		    ControldeDatos.movefirst
		    if IndiceLista>0 then
		      ControldeDatos.MoveTo(IndiceLista+1)
		      ControldeDatos.recordset.DeleteRecord
		    else
		      'Si solo queda un registro, no funciona con el método anterior, ejecutamos una SQL a pelo
		      ControldeDatos.database.SQLExecute "DELETE FROM TAREAS;"
		    end if
		    If ControldeDatos.database.Error Then
		      MsgBox("DB Error: " + ControldeDatos.database.ErrorMessage)
		    else
		      Lista1.RemoveRow(IndiceLista)
		    end if
		    Datacontrol1.RunQuery
		    
		  end select
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ValidarCampo_NO(TipoCampo as String, TipoFormato as String, Campo as Variant) As Variant
		  
		  
		  'dim a1 as integer= val(mid(tHora.text,1,2))
		  'dim a2 as integer=val(mid(tHora.text,4,2))
		  'dim b as string=Textfield5.text
		  'tHora.Text= str(format(a1, "00")) + ":" + str(format(a2, "00"))
		  
		  if TipoCampo="Texto" and TipoFormato="HoraCorta" then
		    if Campo.IsNull or trim(Campo)="" then
		      'Campo="00:00"
		    else
		      static PosicionDosPuntos as integer
		      PosicionDosPuntos=Instr(Campo,":")
		      if PosicionDosPuntos=0 then
		        Campo=str(Campo,"00")+":"+"00"
		      else
		        static LongitudTexto as integer
		        static CaracteresIzqDosPuntos as integer, TextoDerechaPuntos as string
		        'static CaracteresDerDosPuntos as integer
		        
		        LongitudTexto=Len(Campo)
		        TextoDerechaPuntos=Right(Campo,len(Campo)-PosicionDosPuntos)
		        
		        CaracteresIzqDosPuntos=LongitudTexto-(Len(TextoDerechaPuntos)+1)
		        
		        static Hora as string, Minuto as String
		        Hora=mid(Campo,1,CaracteresIzqDosPuntos)
		        Minuto=mid(Campo,CaracteresIzqDosPuntos+2,LongitudTexto-CaracteresIzqDosPuntos-1)
		        Campo=str(Hora,"00")+":" + str(Minuto,"00")
		        'Campo = str(Campo,"00\:00")
		      end if
		    end if
		  end if
		  
		  Return Campo
		  
		End Function
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
