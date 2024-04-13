#tag Class
Protected Class ChatGPTConnection
Inherits URLConnection
	#tag Event
		Sub ContentReceived(URL As String, HTTPStatus As Integer, content As String)
		  //Empty the buffer so we are readt for the next response
		  DataBuffer = ""
		  
		  //Hide the ChatProgressBar on the parent window
		  Me.Parent.ChatProgressBar.Visible = False
		  
		  
		  If HTTPStatus = 200 Then
		    If content.IndexOf("data: [DONE]") <> - 1 Then
		      //OpenAI has finished and the answer can be added to the ContextHistory
		      Var Message As New JSONItem
		      //Add the reply message
		      Message.Value("role") = "assistant"
		      Message.Value("content") = CurrentResponse
		      ContextHistory.Add(message)
		      CurrentResponse = ""
		    End If
		  Else
		    
		    //Check for error
		    Try
		      Var RootDict As Dictionary = ParseJSON(content)
		      Var ErrorDict As Dictionary = RootDict.Value("error")
		      
		      MessageBox(ErrorDict.Lookup("message", "Default error message"))
		      
		    Catch e As JSONException
		      System.DebugLog(e.Message + EndOfLine + content)
		    End Try
		  End if
		  
		  
		  
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Error(e As RuntimeException)
		  MessageBox(e.message + " Error No.: " + e.ErrorNumber.ToString)
		  
		  ContextHistory.RemoveAt(ContextHistory.LastRowIndex)
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub ReceivingProgressed(bytesReceived As Int64, totalBytes As Int64, newData As String)
		  'System.DebugLog(newData)
		  GetJSON(newData)
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub Constructor(Parent As ChatWindow)
		  ParentRef = New WeakRef(Parent)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub GetJSON(DataChunk As String)
		  // Accumulate the incoming data
		  DataBuffer = DataBuffer + DataChunk
		  
		  Var StartMarker As String = "{""id"""
		  Var EndMarker As String = ":null}]}"
		  Var StartPosition, TempEndPosition, EndPosition As Integer
		  
		  // Process complete messages in the DataBuffer
		  While True
		    StartPosition = DataBuffer.IndexOf(StartMarker)
		    If StartPosition = -1 Then Exit // No more complete messages
		    
		    // Create a temporary string that excludes already checked part of the DataBuffer
		    Var TempBuffer As String = DataBuffer.Middle(StartPosition)
		    
		    TempEndPosition = TempBuffer.IndexOf(EndMarker)
		    If TempEndPosition = -1 Then Exit // No valid end marker in the temp buffer, wait for more data
		    
		    // Adjust the EndPosition back to the original buffer's coordinate system
		    EndPosition = StartPosition + TempEndPosition + EndMarker.Length
		    
		    // Extract the JSON message
		    Var JSONMessage As String = DataBuffer.Middle(StartPosition, EndPosition - StartPosition)
		    'System.DebugLog(JSONMessage)
		    // Process the JSON message
		    Me.PopulateOutput(JSONMessage)
		    
		    // Remove the processed message from the DataBuffer
		    DataBuffer = DataBuffer.Middle(EndPosition)
		    
		  Wend
		  
		  '//Empty the buffer if we've received a complete response
		  'If DataBuffer.IndexOf("data: [DONE]") <> - 1 Then
		  'DataBuffer = ""
		  'End if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Parent() As ChatWindow
		  Var ref As ChatWindow = ChatWindow(ParentRef.Value)
		  Return ref
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub PopulateOutput(s As String)
		  'Try
		  '//Parse and process
		  'Var obj As JSONItem = New JSONItem(s)
		  '// Check if the JSON has the "choices" key and it's an array
		  'If obj.HasKey("choices") Then
		  'Var choices As JSONItem = obj.Value("choices")
		  '// Loop through each choice and get the content
		  'For n As Integer = 0 To choices.Count - 1
		  'Var choice As JSONItem = choices.Child(n)
		  'If choice.HasKey("delta") Then
		  'Var delta As JSONItem = choice.Value("delta")
		  'If delta.HasKey("content") Then
		  '// Append the content to the Conversation
		  'Me.Parent.Conversation.AddText(delta.Value("content"))
		  'CurrentResponse = CurrentResponse + delta.Value("content")
		  'Me.Parent.Conversation.VerticalScrollPosition = 100000000
		  'End If
		  'End If
		  'Next
		  'End If
		  'Catch e As JSONException
		  ''System.DebugLog(e.Message + EndOfLine + s)
		  'End Try
		  
		  Try
		    // Parse the JSON string into a Variant
		    Var parsedJSON As Variant = ParseJSON(s)
		    
		    // Attempt to treat the parsed JSON as a Dictionary
		    Var rootDict As Dictionary = parsedJSON
		    
		    If rootDict <> Nil And rootDict.HasKey("choices") Then
		      Var choicesVariant As Variant = rootDict.Value("choices")
		      If choicesVariant <> Nil And choicesVariant.IsArray Then
		        Var choices() As Variant = choicesVariant
		        
		        For Each choiceVariant As Variant In choices
		          Var choiceDict As Dictionary = choiceVariant
		          If choiceDict <> Nil And choiceDict.HasKey("delta") Then
		            Var deltaVariant As Variant = choiceDict.Value("delta")
		            If deltaVariant <> Nil Then
		              Var deltaDict As Dictionary = deltaVariant
		              If deltaDict.HasKey("content") Then
		                Var content As String = deltaDict.Value("content")
		                
		                // Append the content to the Conversation
		                Me.Parent.Conversation.AddText(content)
		                CurrentResponse = CurrentResponse + content
		                Me.Parent.Conversation.VerticalScrollPosition = 100000000
		              End If
		            End If
		          End If
		        Next
		      End If
		    End If
		  Catch e As JSONException
		    // Log the error message
		    'System.DebugLog(e.Message + EndOfLine + s)
		  End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendMessage(prompt As String, optional maintainContext As Boolean, WantedAssistant As String = "")
		  If Not maintainContext Or ContextHistory = Nil Then
		    ContextHistory = New JSONItem
		  End If
		  
		  'Get the connection ready to post
		  RequestHeader("Content-Type") = "application/json"
		  RequestHeader("Authorization") = "Bearer " + Preferences.ApiKey
		  
		  Var j As New JSONItem
		  j.Value("model") = Model
		  'Define the response format "text" or "json_object"
		  Var ResponseFormat As New JSONItem
		  ResponseFormat.Value("type") = "text"
		  j.Value("response_format") = ResponseFormat
		  
		  'Set the temperature (0 to 2 - amount of creativity/hallucination)
		  j.Value("temperature") = Temperature
		  'Set the stream value
		  j.Value("stream") = true
		  
		  
		  'Define the Assistant if first message
		  if WantedAssistant <> "" Then
		    Var Assistant As New JSONItem
		    Assistant.Value("role") = "system"
		    Assistant.Value("content") = WantedAssistant
		    ContextHistory.Add(Assistant)
		  End if
		  
		  'Create a new message
		  Var Message As New JSONItem
		  
		  'Create a message from the prompt
		  Message.Value("role") = "user"
		  Message.Value("content") = prompt
		  ContextHistory.Add(message)
		  
		  'If necessary, trim the context so we don't send one that is too big
		  TrimContext
		  
		  'Add all messages
		  j.Value("messages") = ContextHistory
		  
		  SetRequestContent(j.ToString, "application/json")
		  System.DebugLog(j.ToString)
		  
		  Me.Send("POST", "https://api.openai.com/v1/chat/completions", TimeOut)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TotalContextTokens() As Integer
		  ''Count how many tokens have been used in all the prompts and answers in Messages
		  'Var totalTokens As Integer
		  'For i As Integer = 0 To ContextHistory.LastRowIndex
		  'Var message As JSONItem = ContextHistory.ValueAt(i)
		  'If message.HasKey("usage") Then
		  'Var usage As JSONItem = message.Value("usage")
		  'totalTokens = TotalTokens + usage.Value("total_tokens")
		  'End If
		  'Next
		  '
		  'Return totalTokens
		  
		  //We don't get "usage" when streaming – this is just a temp fix
		  Var MessageLength As  Integer
		  MessageLength = ContextHistory.ToString.Length
		  Return MessageLength
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub TrimContext()
		  'Makes sure that the total tokens in Messages is no larger 90% of the MaximumTokensPerContext.
		  'If it is, it trims the oldest messages from the Messages property until it's below 90%.
		  Var totalTokens As Integer = TotalContextTokens
		  
		  Var percentage As Double = totalTokens / MaximumTokensPerContext
		  If percentage > 0.9 Then '90%
		    'Calculate the maximum number of tokens we should allow to provide enough
		    'room for the new request being added since we can't be sure how many tokens it is
		    Var maxPercentageTokens As Integer = totalTokens / percentage
		    
		    'now remove messages until we are at or below the target number of characters
		    Var tct As Integer
		    Do
		      ContextHistory.RemoveAt(0)
		    Loop Until(maxPercentageTokens <= TotalContextTokens)
		  End If
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private ContextHistory As JSONItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private CurrentResponse As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private DataBuffer As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private MaximumTokensPerContext As Integer = 16385
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Model As String = "gpt-3.5-turbo-0125"
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ParentRef As WeakRef
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Temperature As Double = 0.5
	#tag EndProperty

	#tag Property, Flags = &h21
		Private TimeOut As Integer = 30
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
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
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
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
		#tag ViewProperty
			Name="AllowCertificateValidation"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HTTPStatusCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
