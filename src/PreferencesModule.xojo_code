#tag Module
Protected Module PreferencesModule
	#tag Method, Flags = &h1
		Protected Sub Initialize(appName As String)
		  Preferences = New Prefs(appName)
		End Sub
	#tag EndMethod


	#tag Note, Name = Documentation
		https://blog.xojo.com/2014/01/27/saving-preferences/
		
		Manages your local application preferences.  Preferences are loaded and saved
		automatically.
		
		Usage:
		
		In your App.Open event, call PreferencesModule.Initialize("MyAppName")
		before using the PreferencesModule
		
		Preferences.Load loads your preferences
		Preferences.Save saves your preferences
		
		To assign preferences values:
		
		Preferences.PrefName = value Or Preferences.Set("PrefName") = value
		value = Preferences.PrefName Or value = Preferences.Get("PrefName")
		
		If you attempt to access a preference that does not exist, a
		PreferenceNotFoundException is raised.
		
	#tag EndNote


	#tag Property, Flags = &h0
		Preferences As Prefs
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
