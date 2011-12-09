#Region " Imports "

Imports System.Text
Imports System.Xml

#End Region

''' <summary>
''' A class to hold email signatures for storage in the Sigs list and
''' randomly writing to the outlook sigs file.
''' </summary>
''' <remarks>
''' Written by bjohns - 5/10/2007 2:38:38 PM
''' </remarks>
Public Class Signature

#Region " Properties "

	''' <summary>
	''' Gets or sets the sigs.
	''' </summary>
	''' <value>The sigs.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property Sigs() As List(Of Signature)

	''' <summary>
	''' Gets or sets the sig file.
	''' </summary>
	''' <value>The sig file.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property SigFile() As String

	''' <summary>
	''' Gets or sets the quote file.
	''' </summary>
	''' <value>The quote file.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property QuoteFile() As String

	''' <summary>
	''' Gets or sets the text.
	''' </summary>
	''' <value>The text.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Text() As String

	''' <summary>
	''' Gets or sets the preamble.
	''' </summary>
	''' <value>The preamble.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Preamble() As String

	''' <summary>
	''' Gets or sets the Boilerplate.
	''' </summary>
	''' <value>The Boilerplate.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Boilerplate() As String

#End Region

#Region " constructors "

	''' <summary>
	''' Initializes a new instance of the <see cref="Signature" /> class.
	''' </summary>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub New()
		Sigs = New List(Of Signature)
	End Sub

	''' <summary>
	''' Initializes a new instance of the <see cref="Signature" /> class.
	''' </summary>
	''' <param name="text">The text.</param>
	''' <param name="preamble">The preamble.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:01 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub New(ByVal text As String, ByVal preamble As String)
		Me.Text = text
		Me.Preamble = preamble
	End Sub

	''' <summary>
	''' Initializes a new instance of the <see cref="Signature" /> class.
	''' </summary>
	''' <param name="text">The text.</param>
	''' <param name="preamble">The preamble.</param>
	''' <param name="boilerplate">The Boilerplate.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:01 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub New(ByVal text As String, ByVal preamble As String, ByVal boilerplate As String)
		Me.Text = text
		Me.Preamble = preamble
		Me.Boilerplate = boilerplate
	End Sub

#End Region

#Region " Methods "

	''' <summary>
	''' Chooses the sig.
	''' </summary>
	''' <returns></returns>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Function ChooseSig() As Signature
		If Sigs.Count > 0 Then
			Dim sig As Signature
			Dim r As New Random(DateTime.Now.Millisecond)
			Dim i As Integer = r.Next(0, Sigs.Count - 1)
			sig = New Signature(Sigs.Item(i).Text, Sigs.Item(i).Preamble, Sigs.Item(i).Boilerplate)
			If sig.Text.Length > 0 Then
				Return sig
			End If
		End If
		' if we get this far something is wrong, not sure what
		Throw New Exception("Unable to get a signature.")
	End Function

	''' <summary>
	''' Chooses the sig.
	''' </summary>
	''' <param name="retVal">The ret val.</param>
	''' <returns></returns>
	''' <remarks>
	''' Created: 10/8/2008 at 3:01 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Function ChooseSig(ByRef retVal As Integer) As Signature
		Dim sig As Signature
		If Sigs.Count > 0 Then
			Dim r As New Random(DateTime.Now.Millisecond)
			Dim i As Integer = r.Next(0, Sigs.Count - 1)
			retVal = i
			With Sigs.Item(i)
				sig = New Signature(.Text, .Preamble, .Boilerplate)
			End With
			If sig.Text.Length > 0 Then Return sig
		Else
			GetSigs(Sigs)
			Return ChooseSig()
		End If
		' again, if we get this far something is wrong.
		Throw New Exception("Something is wrong.  Unable to choose a signature.")
	End Function

	''' <summary>
	''' Builds the preamble.
	''' </summary>
	''' <param name="rdr">The RDR.</param>
	''' <param name="sigPreamble">The preamble.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Protected Shared Sub BuildPreamble(ByRef rdr As XmlReader, ByRef sigPreamble As String)
		Dim sb As New StringBuilder
		Dim str As String
		With sb
			rdr.Read()
			str = rdr.ReadElementString("Separator")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("Name")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("NickName")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("OtherInfo")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("Email")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("Phone")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			str = rdr.ReadElementString("WebPage")
			If str.Length > 0 Then
				.Append(str)
				.Append(vbCrLf)
			End If
			.Append(vbCrLf)
			sigPreamble = .ToString
		End With
	End Sub

	''' <summary>
	''' Gets the Boilerplate.
	''' </summary>
	''' <param name="rdr">The RDR.</param>
	''' <returns></returns>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Protected Shared Function GetBoilerplate(ByRef rdr As XmlReader) As String
		Dim sb As New StringBuilder
		Dim bp As String
		bp = rdr.ReadElementString("Boilerplate")
		If Not bp = String.Empty Then
			With sb
				.Append(vbCrLf & vbCrLf)
				.Append(bp)
			End With
			Return sb.ToString
		End If
		Return bp
	End Function

	''' <summary>
	''' Gets the sigs.
	''' </summary>
	''' <param name="signatures">The sigs.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Sub GetSigs(ByRef signatures As List(Of Signature))
		Dim sigPreamble As String = String.Empty
		If signatures Is Nothing Then
			signatures = New List(Of Signature)
		End If
		Dim xmlSettings As New XmlReaderSettings
		xmlSettings.IgnoreComments = True
		xmlSettings.IgnoreWhitespace = True
		Try
			Using xmlRdr As XmlReader = XmlReader.Create(QuoteFile, xmlSettings)
				With xmlRdr
					.Read()
					.Read()
					BuildPreamble(xmlRdr, sigPreamble)
					Dim sigBoilerplate As String = GetBoilerplate(xmlRdr)
					.Read()
					While .Read
						Dim s As New Signature(.Value, sigPreamble, sigBoilerplate)
						If s.Text.Length > 0 Then
							signatures.Add(s)
						End If
					End While
				End With
				xmlRdr.Close()
			End Using
		Catch ex As IO.FileNotFoundException
			Throw New QuoteFileNotFoundException(ex, QuoteFile)
		Catch ex As Exception
			Throw New Exception(String.Format("An error occurred trying to get the signatures from the file ""{0}""", QuoteFile), ex)
		End Try
	End Sub

	''' <summary>
	''' Returns a <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
	''' </summary>
	''' <returns>
	''' A <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
	''' </returns>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Public Overrides Function ToString() As String
		Dim sb As New StringBuilder
		With sb
			.Append(Preamble)
			.Append(vbCrLf)
			.Append(Text)
			.Append(vbCrLf)
			.Append(Boilerplate)
		End With
		Return sb.ToString
	End Function

#End Region

End Class

''' <summary>
''' An exception to throw if the quote file isn't found.
''' </summary>
''' <remarks>
''' Created: 10/8/2008 at 3:03 PM
''' By: bjohns.
''' </remarks>
Public Class QuoteFileNotFoundException
	Inherits Exception

	''' <summary>
	''' Initializes a new instance of the <see cref="QuoteFileNotFoundException" /> class.
	''' </summary>
	''' <param name="ex">The ex.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:03 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub New(ByVal ex As Exception, ByVal quoteFile As String)
		MyBase.New("The specified quote file " & quoteFile & " was not found.", ex)
	End Sub

End Class