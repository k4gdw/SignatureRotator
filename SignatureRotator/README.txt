K4GDW Signature Rotator ReadMe

bjohns - 10/8/2008 9:43:40 AM
Updated it to DotNet 2.0 and created a Live Writer plugin.  To install the
Live Writer plugin, extract the LiveWriterRandomSigPlugin.dll into your
Live Writer Plugins directory, usually:
"c:\Program Files\Windows Live\Writer\Plugins"  Also, you'll need to edit
the WindowsLiveWriter.exe.config file to make it look like this:

<configuration>
	<appSettings>
		<add	key="SigfileLocation" 
				value="Path to signature file that is rotated by the SigRotator program."/>
	</appSettings>
	<system.diagnostics>
		<sources>
			<!-- This section defines the logging configuration for My.Application.Log -->
			<source name="DefaultSource"
					switchName="DefaultSwitch">
				<listeners>
					<add name="SignatureRotatorWLPlugin"/>
				</listeners>
			</source>
		</sources>
		<switches>
			<add name="DefaultSwitch" value="Information" />
		</switches>
		<sharedListeners>
			<add name="SignatureRotatorWLPlugin"
				 type="System.Diagnostics.EventLogTraceListener"
				 initializeData="SignatureRotatorWLPlugin"/>
		</sharedListeners>
	</system.diagnostics>
	<startup>
    	<supportedRuntime version="v2.0.50727"/>
	</startup>
	<runtime>
		<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
			<probing PrivatePath="Plugins"/>
			<dependentAssembly>
				<assemblyIdentity
					name="WindowsLive.Writer.Api"
					publicKeyToken="31BF3856AD364E35" />
				<bindingRedirect oldVersion="1.0.0.0"
					newVersion="1.1.0.0" />
			</dependentAssembly>
		</assemblyBinding>
	</runtime>
</configuration>

bjohns - 6/28/2007 11:38:13 AM
Added support for a "disclaimer".  This field in the myquotes.xml file is
intended to allow the user to put in some boiler plate info that they wish
to appear on all emails.  For example, at work I use it to display the
suggested confidentiality notice from our IT security  folks.

Written By:  Bryan Johns - 5/11/2007 2:00:41 PM

Thank you for checking out this program.  It's the result of some significant
frustration with Outlook's lack of a similar built in feature.  Heck, even
PINE, the standard email client for the *NIX command shell, with the help of
a little shell script, can do this natively.

Anyway, to use this program you must have at least the .net 2.0 framework
already installed.  If you don't have this, which should be very rare by now,
you can download it from Microsoft at this URL:

http://www.microsoft.com/downloads/details.aspx?familyid=0856EACB-4362-4B0D
-8EDD-AAB15C5E04F5&displaylang=en

You'll need to make sure that the URL is on a single line.

Then, extract the contents of this zip file into a directory of your choice.

Edit the SignatureRotator.exe.config file in your favorite text editor,
following the guidelines in the comments contained therin.

Then, edit the myquotes.xml file with a text editor.  The contents should
be rather self explanatory.  Replace the text in the various attributes
with what you want displayed.  For example, if you edit the file to
look like this:

<?xml version="1.0" encoding="utf-8"?>
<Signature>
	<Separator>--</Separator>
	<Name>John Doe</Name>
	<NickName>Johnny B Good</NickName>
	<OtherInfo></OtherInfo>
	<Email>jbgood@somewhere.com</Email>
	<Phone></Phone>
	<WebPage>http://www.jbgood.com</WebPage>
	<Disclaimer></Disclaimer>
	<Quotes>
		...
	</Quotes>
</Signature>

The signature saved for Outlook, or whatever, to use will be like this:

--
John Doe
Johnny B Good
jbgood@somewhere.com
http://www.jbgood.com

with a random quote two lines below the web address.  The system ignores
blank values, so if you only want the quote to be listed, leave everything
other than the Separator entry blank.  Another option, if you want a 
different quote format, you can try playing around with leaving everything
but the <OtherInfo></OtherInfo> attribute blank and put whatever you want
there.  It'll take some experimenting to get it right but feel free to do
so if you wish.  You can probably do the same thing with multiple different
signature formats by using the <Quote></Quote> attributes.  That could
allow you to have multiple totally unique signatures.

whatever text you put in the Disclaimer field should appear here.