#Region " Imports "

Imports System
Imports System.IO
Imports System.Configuration
Imports System.Windows.Forms
Imports WindowsLive.Writer.Api

#End Region

''' <summary>
''' This is a Windows Live Writer plugin class to read in the current signature
''' from the SignatureRotator application into a blog post.
''' </summary>
''' <remarks>
''' It depends on the SignatureRotator application to rotate signatures.
''' </remarks>
<WriterPlugin("A00D1D53-9111-4d5a-8D0D-75FB05423DC2", "Random Signature", _
	 Description:="Inserts a signature block into the blog post", _
	 HasEditableOptions:=False, ImagePath:="gdw16.ico", _
	 PublisherURL:="http://www.greendragonweb.com")> _
<InsertableContentSource("Random Signature")> _
Public Class SignatureRotatorPlugin
	Inherits ContentSource

	''' <summary>
	''' Creates the content.
	''' </summary>
	''' <param name="dialogOwner">The dialog owner.</param>
	''' <param name="content">The content.</param>
	''' <returns></returns>
	Public Overrides Function CreateContent(ByVal dialogOwner As IWin32Window, ByRef content As String) As DialogResult
		content = Signature()
		Return DialogResult.OK
	End Function

	''' <summary>
	''' reads the signature file specified in the app.config.
	''' </summary>
	''' <returns>
	''' The contents of the signature file.
	''' </returns>
	''' <remarks></remarks>
	Private Function Signature() As String
		Dim result As String
		Dim sigfile As String = ConfigurationManager.AppSettings("SigfileLocation")
		Using sr As StreamReader = File.OpenText(sigfile)
			result = sr.ReadToEnd
			sr.Close()
			sr.Dispose()
		End Using
		CleanString(result)
		Return result
	End Function

	''' <summary>
	''' Cleans the string.
	''' </summary>
	''' <param name="text">The text.</param>
	Private Sub CleanString(ByRef text As String)
		text = text.Replace(vbCrLf, "<br />")
		text = text.Replace("""", "&quot;")
	End Sub

End Class
