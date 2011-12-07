#Region " Imports "

Imports System.Text
Imports System.Xml
Imports System.Collections

#End Region

''' <summary>
''' A class to hold email signatures for storage in the Sigs list and
''' randomly writing to the outlook sigs file.
''' </summary>
''' <remarks>
''' Written by bjohns - 5/10/2007 2:38:38 PM
''' </remarks>
Public Class Signature
	Implements IDisposable

#Region " Properties "

	Protected Shared _sigs As Generic.List(Of Signature)
	Protected Shared _sigFile As String
	Protected Shared _quoteFile As String
	Private _text As String = String.Empty
	Private _preamble As String = String.Empty
	Private _disclaimer As String = String.Empty

	''' <summary>
	''' Gets or sets the sigs.
	''' </summary>
	''' <value>The sigs.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property Sigs() As Generic.List(Of Signature)
		Get
			Return _sigs
		End Get
		Set(ByVal value As Generic.List(Of Signature))
			_sigs = value
		End Set
	End Property

	''' <summary>
	''' Gets or sets the sig file.
	''' </summary>
	''' <value>The sig file.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property SigFile() As String
		Get
			Return _sigFile
		End Get
		Set(ByVal value As String)
			_sigFile = value
		End Set
	End Property

	''' <summary>
	''' Gets or sets the quote file.
	''' </summary>
	''' <value>The quote file.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 2:59 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Property QuoteFile() As String
		Get
			Return _quoteFile
		End Get
		Set(ByVal value As String)
			_quoteFile = value
		End Set
	End Property

	''' <summary>
	''' Gets or sets the text.
	''' </summary>
	''' <value>The text.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Text() As String
		Get
			Return _text
		End Get
		Set(ByVal value As String)
			_text = value
		End Set
	End Property

	''' <summary>
	''' Gets or sets the preamble.
	''' </summary>
	''' <value>The preamble.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Preamble() As String
		Get
			Return _preamble
		End Get
		Set(ByVal value As String)
			_preamble = value
		End Set
	End Property

	''' <summary>
	''' Gets or sets the disclaimer.
	''' </summary>
	''' <value>The disclaimer.</value>
	''' <remarks>
	''' Created: 10/8/2008 at 3:00 PM
	''' By: bjohns.
	''' </remarks>
	Public Property Disclaimer() As String
		Get
			Return _disclaimer
		End Get
		Set(ByVal value As String)
			_disclaimer = value
		End Set
	End Property

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
		_text = text
		_preamble = preamble
	End Sub

	''' <summary>
	''' Initializes a new instance of the <see cref="Signature" /> class.
	''' </summary>
	''' <param name="text">The text.</param>
	''' <param name="preamble">The preamble.</param>
	''' <param name="disclaimer">The disclaimer.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:01 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub New(ByVal text As String, ByVal preamble As String, ByVal disclaimer As String)
		_text = text
		_preamble = preamble
		_disclaimer = disclaimer
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
		Dim sig As Signature
		If Sigs.Count > 0 Then
			Dim r As New Random(System.DateTime.Now.Millisecond)
			Dim i As Integer = r.Next(0, Sigs.Count - 1)
			sig = New Signature(Sigs.Item(i).Text, Sigs.Item(i).Preamble, Sigs.Item(i).Disclaimer)
			If sig.Text.Length > 0 Then
				Return sig
			Else
				Return ChooseSig()
			End If
		Else
			GetSigs(Sigs)
			Return ChooseSig()
		End If
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
			Dim r As New Random(System.DateTime.Now.Millisecond)
			Dim i As Integer = r.Next(0, Sigs.Count - 1)
			retVal = i
			With Sigs.Item(i)
				sig = New Signature(.Text, .Preamble, .Disclaimer)
			End With
			If sig.Text.Length > 0 Then
				Return sig
			Else
				Return ChooseSig()
			End If
		Else
			GetSigs(Sigs)
			Return ChooseSig()
		End If
	End Function

	''' <summary>
	''' Builds the preamble.
	''' </summary>
	''' <param name="rdr">The RDR.</param>
	''' <param name="preamble">The preamble.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Protected Shared Sub BuildPreamble(ByRef rdr As XmlTextReader, ByRef preamble As String)
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
			preamble = .ToString
		End With
		sb = Nothing
	End Sub

	''' <summary>
	''' Gets the disclaimer.
	''' </summary>
	''' <param name="rdr">The RDR.</param>
	''' <returns></returns>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Protected Shared Function GetDisclaimer(ByRef rdr As XmlTextReader) As String
		Dim sb As New StringBuilder
		'rdr.Read()
		Try
			Dim dis As String = String.Empty
			dis = rdr.ReadElementString("Disclaimer")
			If Not dis = String.Empty Then
				With sb
					.Append(vbCrLf & vbCrLf)
					.Append(dis)
				End With
				Return sb.ToString
			Else
				Return dis
			End If
		Finally
			sb = Nothing
		End Try
	End Function

	''' <summary>
	''' Gets the sigs.
	''' </summary>
	''' <param name="sigs">The sigs.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:02 PM
	''' By: bjohns.
	''' </remarks>
	Public Shared Sub GetSigs(ByRef sigs As Generic.List(Of Signature))
		Dim preamble As String = String.Empty
		If sigs Is Nothing Then
			sigs = New Generic.List(Of Signature)
		End If
		Using xmlRdr As New XmlTextReader(_quoteFile)
			Try
				With xmlRdr
					.WhitespaceHandling = WhitespaceHandling.None
					.Read()
					.Read()
					BuildPreamble(xmlRdr, preamble)
					Dim disclaimer As String = GetDisclaimer(xmlRdr)
					.Read()
					While .Read
						Dim s As New Signature(.Value, preamble, disclaimer)
						If s.Text.Length > 0 Then
							sigs.Add(s)
						End If
					End While
				End With
			Catch ex As System.IO.FileNotFoundException
				Throw New QuoteFileNotFoundException(ex, _quoteFile)
			Catch ex As Exception
				Throw ex
			Finally
				xmlRdr.Close()
			End Try
		End Using
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
		Dim sb As New System.Text.StringBuilder
		With sb
			.Append(Me.Preamble)
			.Append(vbCrLf)
			.Append(Me.Text)
			.Append(vbCrLf)
			.Append(Me.Disclaimer)
			Dim retval As String = sb.ToString
			sb = Nothing
			Return retval
		End With
	End Function

#End Region

#Region " IDisposable Support "

	Private disposedValue As Boolean = False		' To detect redundant calls

	''' <summary>
	''' Releases unmanaged and - optionally - managed resources
	''' </summary>
	''' <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
	''' <remarks>
	''' Created: 10/8/2008 at 3:37 PM
	''' By: bjohns.
	''' </remarks>
	Protected Overridable Sub Dispose(ByVal disposing As Boolean)
		If Not Me.disposedValue Then
			If disposing Then
				' TODO: free other state (managed objects).
			End If

			' TODO: free your own state (unmanaged objects).
			' TODO: set large fields to null.
		End If
		Me.disposedValue = True
	End Sub

	''' <summary>
	''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
	''' </summary>
	''' <remarks>
	''' Created: 10/8/2008 at 3:37 PM
	''' By: bjohns.
	''' </remarks>
	Public Sub Dispose() Implements IDisposable.Dispose
		' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub

	''' <summary>
	''' Allows an <see cref="T:System.Object" /> to attempt to free resources and perform other cleanup operations before the <see cref="T:System.Object" /> is reclaimed by garbage collection.
	''' </summary>
	''' <remarks>
	''' Created: 10/8/2008 at 3:36 PM
	''' By: bjohns.
	''' </remarks>
	Protected Overrides Sub Finalize()
		Dispose(False)
		MyBase.Finalize()
	End Sub

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
	Inherits ApplicationException

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