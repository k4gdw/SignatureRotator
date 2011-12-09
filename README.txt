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