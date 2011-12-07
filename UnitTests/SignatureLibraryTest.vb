Imports System
Imports MbUnit.Framework
Imports K4GDW.RandomSig.Signature

''' <summary>
''' A test fixture to hold the SignatureLibrary unit tests.
''' </summary>
''' <remarks>Created: 10/7/2008 at 1:24 PM by bjohns.?</remarks>
<TestFixture()> _
Public Class SignatureLibraryTest

	''' <summary>
	''' Sets up the test fixture.
	''' </summary>
	''' <remarks>Created: 10/7/2008 at 1:23 PM by bjohns.?</remarks>
	<SetUp()> _
	Public Sub Setup()
		Sigs = New Generic.List(Of Signature)
		QuoteFile = "myquotes.xml"
		SigFile = "testsig.txt"
		GetSigs(Sigs)
	End Sub

	''' <summary>
	''' Tears down the test fixture.
	''' </summary>
	''' <remarks>Created: 10/7/2008 at 1:23 PM by bjohns.?</remarks>
	<TearDown()> _
	Public Sub TearDown()
		If Sigs.Count > 0 Then
			Sigs.Clear()
		End If
		Sigs = Nothing
	End Sub

	''' <summary>
	''' Tests that the correct sigs are read in from the test quote file.
	''' </summary>
	''' <remarks>Created: 10/7/2008 at 1:23 PM by bjohns.</remarks>
	<Test()> _
	Public Sub TestGetSigs()
		Assert.IsTrue(Sigs.Count = 4, "The test collection didn't contain the correct number of signatures.")
		For x As Integer = 0 To 3 Step 1
			Assert.IsTrue(Sigs(x).Text = "Test Quote " & CStr(x + 1), "Test Quote " & CStr(x + 1) & " didn't load.")
		Next
	End Sub

	''' <summary>
	''' Tests that a "QuoteFileNotFoundException" is thrown
	''' if a non-existant quote file is specified.
	''' </summary>
	''' <remarks>Created: 10/7/2008 at 1:24 PM by bjohns.</remarks>
	<Test(), ExpectedException(GetType(QuoteFileNotFoundException))> _
	Public Sub TestQuoteFileNotFound()
		QuoteFile = "filenotfound.xml"
		Sigs.Clear()
		GetSigs(Sigs)
		Assert.Fail("Should have triggered a QuoteFileNotFoundException")
	End Sub

	''' <summary>
	''' Tests the ChooseSig() method and it's overload that
	''' accepts an integer to be able to also return the index
	''' of the chosen sig.
	''' </summary>
	''' <remarks>
	''' Created: 7/2/2009 at 12:30 PM
	''' By: bjohns.
	''' </remarks>
	<Test()> _
	Public Sub TestChooseSig()
		Dim i As Integer
		Dim sig As Signature
		sig = ChooseSig()
		Assert.IsNotNull(sig, "The first sig returned null.")
		Dim sig2 As Signature
		sig2 = ChooseSig(i)
		Assert.IsNotNull(sig2, "The second sig returned null.")
		Assert.IsTrue((i <= 3 And i >= 0), "The index returned was out of range.")
	End Sub

	''' <summary>
	''' Tests the RemoveAt method to make sure the signature specified is
	''' removed.
	''' </summary>
	''' <remarks>
	''' Created: 7/2/2009 at 12:31 PM
	''' By: bjohns.
	''' </remarks>
	<Test()> _
	Public Sub TestDropSig()
		Dim i As Integer
		Dim sig As Signature = ChooseSig(i)
		Sigs.RemoveAt(i)
		Assert.IsFalse(Sigs.Contains(sig), "The signature wasn't removed.")
		Assert.IsTrue(Sigs.Count = 3, "There are still 4 signatures.")
	End Sub

	''' <summary>
	''' Tests that an "ArgumentOutOfRangeException" is thrown if
	''' a sig with an index out of range is requested.
	''' </summary>
	''' <remarks>
	''' Created: 7/2/2009 at 12:33 PM
	''' By: bjohns.
	''' </remarks>
	<Test(), ExpectedException(GetType(ArgumentOutOfRangeException))> _
	Public Sub TestListBoundaries()
		Dim sig1 As Signature = Sigs(4)
		Assert.Fail("An out of bounds exception should have been triggered.")
	End Sub

End Class
