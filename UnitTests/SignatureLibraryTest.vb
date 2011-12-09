Imports System
Imports NUnit.Framework
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
        Sigs = New List(Of Signature)
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
        Assert.IsTrue(Sigs.Count = 4, String.Format("There were {0} signatures.", Sigs.Count))
        For x As Integer = 0 To 3 Step 1
            Assert.IsTrue(Sigs(x).Text = "Test Quote " & CStr(x + 1), "Test Quote " & CStr(x + 1) & " didn't load.")
        Next
	End Sub

	''' <summary>
	''' Tests that a "QuoteFileNotFoundException" is thrown
	''' if a non-existant quote file is specified.
	''' </summary>
	''' <remarks>Created: 10/7/2008 at 1:24 PM by bjohns.</remarks>
    <Test(), ExpectedException(GetType(QuoteFileNotFoundException))>
    Public Sub TestQuoteFileNotFound()
        QuoteFile = "filenotfound.xml"
        Sigs.Clear()
        GetSigs(Sigs)
    End Sub

    ''' <summary>
    ''' Tests the throws exception if pointed at non XML file.
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Created:  12/7/2011 at 5:47 PM.<br />
    ''' By: Bryan Johns, K4GDW<br />
    ''' Email:  bjohns@greendragonweb.com
    ''' </para>
    ''' </remarks>
    <Test(), ExpectedException(GetType(Exception))>
    Public Sub TestThrowsExceptionIfPointedAtNonXMLFile()
        QuoteFile = "nonxmlfile.txt"
        Sigs.Clear()
        GetSigs(Sigs)
    End Sub

	<Test()>
	Public Sub An_Empty_Boilerplate_Returns_An_Empty_String()
		QuoteFile = "emptydisclaimerquotes.xml"
		Sigs.Clear()
		GetSigs(Sigs)
		Dim s As Signature = ChooseSig()
		Assert.IsNullOrEmpty(s.Boilerplate)
	End Sub

    <Test()>
    Public Sub After_Choosing_All_Sigs_It_Reloads()
        Dim s As Signature
        For x As Integer = 1 To 6
            Dim i As Integer
            s = ChooseSig(i)
            Sigs.RemoveAt(i)
        Next
        Assert.IsTrue(0 < Sigs.Count < 5, String.Format("there were {0} signatures.", Sigs.Count))
        Assert.IsNotNullOrEmpty(s.Text)
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
        Assert.IsTrue(Sigs.Count = 3, String.Format("There are {0} signatures.", Sigs.Count))
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
        Console.WriteLine(Sigs(5))
    End Sub

    ''' <summary>
    ''' Determines whether this instance [can get to string].
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Created:  12/7/2011 at 5:29 PM.<br />
    ''' By: Bryan Johns, K4GDW<br />
    ''' Email:  bjohns@greendragonweb.com
    ''' </para>
    ''' </remarks>
    <Test()>
    Public Sub CanGetToString()
        Dim sig As Signature = ChooseSig()
        Assert.IsNotNullOrEmpty(sig.ToString())
    End Sub

    <Test()>
    Public Sub CanWorkWithNullSigsList()
        Dim lsigs As List(Of Signature)
        GetSigs(lsigs)
        Assert.IsNotNull(lsigs)
        Assert.IsTrue(lsigs.Count > 0)
    End Sub

End Class
