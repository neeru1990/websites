From ipowerweb.com

A. Please configure the following settings for your script
inorder to correct the problem. Also you can go into your
control panel and go into your file manager. Click on the
paticular file or folder and to the right side you will have
the option to change your permissions. You can enter 755 to
make your file executable.


1. do your editing in notepad and not wordpad

2. make sure the top line says:
#!/usr/bin/perl

3. make sure the @referers line looks like this:
@referers=('your_domain_name','www.your_domain_name');

4. make sure the path to sendmail looks like this:
$mailprog = '/usr/sbin/sendmail';

5. make sure the permissions on the script are set to 755


In your html form make sure the form action looks like this:

<form action="http://your_domain_name.com/cgi-bin/formmail.pl" method=post>

You also need to have a hidden field in your form that looks
like this (replace it with your email address):

<input type=hidden name="recipient" value="email@your.host.com">

If you wish to redirect the user to a different URL, rather than having them see the default response to the fill-out form, you can use this hidden variable to send them to a pre-made HTML page. 

<input type=hidden name="redirect" value="http://your.host.com/to/file.html">