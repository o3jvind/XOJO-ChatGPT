#tag DesktopWindow
Begin DesktopWindow ChatWindow
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   Height          =   528
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   518635519
   MenuBarVisible  =   False
   MinimumHeight   =   528
   MinimumWidth    =   826
   Resizeable      =   True
   Title           =   "Chat GPT Example"
   Type            =   0
   Visible         =   True
   Width           =   826
   Begin DesktopButton SendButton
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "Send"
      Default         =   True
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   726
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   488
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin DesktopTextArea Conversation
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   True
      AllowStyledText =   True
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      HasHorizontalScrollbar=   False
      HasVerticalScrollbar=   True
      Height          =   416
      HideSelection   =   True
      Index           =   -2147483648
      Italic          =   False
      Left            =   0
      LineHeight      =   0.0
      LineSpacing     =   1.0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Multiline       =   True
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   60
      Transparent     =   False
      Underline       =   False
      UnicodeMode     =   1
      ValidationMask  =   ""
      Visible         =   True
      Width           =   826
   End
   Begin DesktopTextField Prompt
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   True
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   2
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   488
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   694
   End
   Begin DesktopPopupMenu AssistantPopUp
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialValue    =   ""
      Italic          =   False
      Left            =   610
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   2
      SelectedRowIndex=   0
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   196
   End
   Begin DesktopProgressBar ChatProgressBar
      Active          =   False
      AllowAutoDeactivate=   True
      AllowTabStop    =   True
      Enabled         =   True
      Height          =   20
      Indeterminate   =   True
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      MaximumValue    =   100
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   20
      Transparent     =   False
      Value           =   0.0
      Visible         =   False
      Width           =   578
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  ChatConnection = New ChatGPTConnection(Me)
		End Sub
	#tag EndEvent


	#tag MenuHandler
		Function FileNewConversation() As Boolean Handles FileNewConversation.Action
		  Var w as new ChatWindow
		  Return True
		  
		End Function
	#tag EndMenuHandler


	#tag Property, Flags = &h21
		Private ChatConnection As ChatGPTConnection
	#tag EndProperty


#tag EndWindowCode

#tag Events SendButton
	#tag Event
		Sub Pressed()
		  //Add what the user typed to the conversation
		  If Conversation.Text <> "" Then Conversation.AddText(EndOfLine + EndOfLine)
		  Conversation.SelectionBold = True
		  Conversation.AddText("You" + EndOfLine)
		  Conversation.SelectionBold = False
		  Conversation.AddText(Prompt.Text)
		  
		  Try
		    //Disconnect the Chatconnection if its already conected
		    ChatConnection.Disconnect
		    
		    Var Assistant As String = ""
		    //Define the assistant if this is the first request
		    If AssistantPopUp.Enabled then
		      Assistant = AssistantPopUp.RowTagAt(AssistantPopUp.SelectedRowIndex)
		    End if
		    
		    //Disable the AssistantPopUp after first run
		    AssistantPopUp.Enabled = False
		    
		    //Send the user's prompt to ChatGPT then clear the Prompt field
		    ChatConnection.SendMessage(Prompt.Text, True, Assistant)
		    Prompt.Text = ""
		    
		    //Make some room and make ready for the response
		    Conversation.AddText(EndOfLine + EndOfLine)
		    Conversation.SelectionBold = True
		    Conversation.AddText(AssistantPopUp.SelectedRowText + EndOfLine)
		    Conversation.SelectionBold = False
		    
		    //Make the ChatProgressBar visible
		    ChatProgressBar.Visible = True
		    
		    
		  Catch e As ChatGPTException
		    System.Beep
		    MessageBox("Error: " + e.Message)
		  End Try
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Conversation
	#tag Event
		Sub Opening()
		  Me.SelectionFontSize = 14
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Prompt
	#tag Event
		Sub TextChanged()
		  SendButton.Enabled = Me.Text <> ""
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events AssistantPopUp
	#tag Event
		Sub Opening()
		  Me.AddRow("ChatGPT")
		  Me.RowTagAt(Me.LastAddedRowIndex) = "You are helpful assistant"
		  
		  Me.AddRow("Grumpy Old Man")
		  Me.RowTagAt(Me.LastAddedRowIndex) = "You're a grumpy old man who isn't willing to help in any way. You tend to believe that everything was better in the old days. You have a strong aversion to capitalized products and company names, but can reluctantly talk about them. You are patronizing, especially if people make typos or spelling mistakes, and don't fail to point out people's errors."
		  
		  Me.AddRow("The Politician")
		  Me.RowTagAt(Me.LastAddedRowIndex) = "You never give concrete answers. You are never specific and question peoples agenda when they ask you something. If someone asks you concrete questions you often reply with hard to answer ethical questions."
		  
		  Me.AddRow("The Yes Man")
		  Me.RowTagAt(Me.LastAddedRowIndex) = "You are a Yes Man. Servile to the extreme. You will do anything to please anybody. You never give advice against the qustions you get."
		  
		  Me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  Me.Tooltip = Me.RowTagAt(Me.SelectedRowIndex)
		  Self.Title = Me.SelectedRowText
		End Sub
	#tag EndEvent
#tag EndEvents
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
		Name="Interfaces"
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
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
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
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
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
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
