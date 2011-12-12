K4GDW Signature Rotator ReadMe


This is a simple little utility that rotates email signatures for email programs
that support an external signature file but do not provide a means of rotating
them automatically, such as Outlook.

For an example of how to configure it, look at the included myquotes.xml and
the app.config in the SignatureRotator project.

In short, you configure your signature and add quotes in myquotes.xml.  That
file name is the default, it is configurable from app.config.  Then you edit
app.config to tell the program where to stash and what to call the generated
signature file.  Then, you start the program.  It runs in your system tray
changing the signature file on a timed basis (30 seconds by default) which
is set in app.config.  The next step is to open your email app and configure
it's custom signature file feature and point it at the generated file.

I've tested this with Outlook 2003 and 2007.  It should work with just about
any windows email program that lets you configure a signature file.

<?xml version="1.0" encoding="utf-8"?>
<Signature>
	<Separator>--</Separator>
	<Name>John Doe</Name>
	<NickName>Johnny B Goode</NickName>
	<OtherInfo></OtherInfo>
	<Email>jbgoode@somewhere.com</Email>
	<Phone></Phone>
	<WebPage>http://www.jbgoode.com</WebPage>
	<Boilerplate></Boilerplate>
	<Quotes>
		<Quote>Some witty words from someone...</Quote>
		...
	</Quotes>
</Signature>

The signature saved for Outlook, or whatever, to use will be like this:

--
John Doe
Johnny B Goode
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

whatever text you put in the Boilerplate field should appear here.