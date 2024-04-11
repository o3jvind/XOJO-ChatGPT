#tag Class
Private Class Prefs
	#tag Method, Flags = &h0
		Sub Constructor(appName As String)
		  mPreferences = New JSONItem
		  mPreferences.Compact = False
		  
		  mAppName = appName
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Get(name As String) As Variant
		  // Allow lookup of preference using this syntax:
		  // Var top As Integer = Preferences.Get("MainWindowTop")
		  
		  If mPreferences <> Nil Then
		    If mPreferences.HasName(name) Then
		      Return mPreferences.Value(name)
		    Else
		      Raise New PreferenceNotFoundException(name)
		    End If
		  Else
		    Return -1
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Load() As Boolean
		  Var prefFolder As FolderItem = SpecialFolder.ApplicationData.Child(mAppName)
		  If Not prefFolder.Exists Then
		    prefFolder.CreateAsFolder
		  End If
		  
		  Var prefName As String = mAppName + ".pref"
		  
		  mPreferencesFile = prefFolder.Child(prefName)
		  
		  If mPreferencesFile.Exists Then
		    
		    Var input As TextInputStream
		    input = TextInputStream.Open(mPreferencesFile)
		    
		    Var data As String = input.ReadAll
		    input.Close
		    
		    Try
		      mPreferences.Load(data)
		      
		      Return True
		    Catch e As JSONException
		      Return False
		    End Try
		  Else
		    Return False
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Operator_Lookup(name As String) As Variant
		  // Allow lookup of preference using this syntax:
		  // Var top As Integer = Preferences.MainWindowTop
		  
		  Return Get(name)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(name As String, Assigns value As Variant)
		  // Set a preference using this syntax:
		  // Preferences.MainWindowTop = 345
		  
		  Set(name) = value
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Save() As Boolean
		  If mPreferencesFile <> Nil Then
		    Var data As String = mPreferences.ToString
		    
		    Try
		      Var output As TextOutputStream
		      output = TextOutputStream.Create(mPreferencesFile)
		      
		      output.Write(data)
		      
		      Return True
		    Catch e As IOException
		      Return False
		    End Try
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Set(name As String, Assigns value As Variant)
		  If mPreferences <> Nil Then
		    // Set a preference using this syntax:
		    // Preferences.Set("MainWindowTop") = 345
		    
		    mPreferences.Value(name) = value
		  End If
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mAppName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		mPreferences As JSONItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPreferencesFile As FolderItem
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
End Class
#tag EndClass
