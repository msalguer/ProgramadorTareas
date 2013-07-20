#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Open()
		  
		  
		End Sub
	#tag EndEvent


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Borrar", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Borrar"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Borrar"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Salir", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Salir"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
