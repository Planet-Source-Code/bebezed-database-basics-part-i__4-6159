<div align="center">

## Database Basics: Part I


</div>

### Description

Since ASP is especially good at reading and writing to databases, let's start with a very simple database and scripts that we'll eventually build into a guestbook...

Reprinted with permission from http://www.web-savant.com/users/kathi/asp
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BebeZed](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bebezed.md)
**Level**          |Advanced
**User Rating**    |4.7 (42 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bebezed-database-basics-part-i__4-6159/archive/master.zip)





### Source Code

<font face="Verdana" size="2">For now, start with a database with Name, City,
State, and Country. I've built my database with MS Access, but you can use any
ODBC-compliant database you'd like. First, follow the link below to view the ASP
page, then click on "view code" to see the code behind the page. If
you'd like, you can <a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample.zip">download</a>
the MS Access database and scripts. (Zip format, 18K)</font>
<p><a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample1.asp" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/sample1code.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a></font>
<p><a name="modify"><b><font face="Verdana" size="2">Modifying Records</font></b></a><font face="Verdana" size="2"><br>
Now that you've seen how simple it is to display contents of your database,
let's go to the next step - modifying the records that exist in the database.
We'll start with a very simple page that queries the database and displays a
summary of the contents (in this case, we'll just list the names), and then
allows you to click on a listing to modify the information contained in that
record.</font>
<p><a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample2.asp" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/sample2code.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a>      <a href="#top"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</a></font>
<p><a name="addnew"><b><font face="Verdana" size="2">Adding New Records</font></b></a><font face="Verdana" size="2"><br>
The above works just fine if you have records in the database. But what if you
want to add records? No problem! First we'll create a regular html form that
passes the variables to a script for processing, then the script itself that
adds the new record to the database. In this script, we'll also introduce simple
form validation prior to processing.</font>
<p><a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample3.html" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/sample3code.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a>      <a href="#top"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</a></font>
<p><a name="delete"><b><font face="Verdana" size="2">Deleting Records</font></b></a><font face="Verdana" size="2"><br>
The last thing we'll need to do is add a provision to delete unwanted records
from our database. Please note that in the example, all records can be deleted
except for one - I wanted to make sure that there would be at least one record
in the database at all times for demonstration purposes.</font>
<p><a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample4.asp" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/sample4code.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a>      <a href="#top"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</a></font>
<p><a name="selfsubmit"><b><font face="Verdana" size="2">Submitting Scripts to
Themselves</font></b></a><font face="Verdana" size="2"><br>
Sometimes its advantageous to use a single script for multiple purposes. But how
do we accomplish this? By using scripts that submit back to themselves, using
some sort of "flag" to determine what actions need to be performed.
Let's use our database scripts as an example. We had an html form for entering
information into a new record, which submits to a separate ASP script to
actually insert the record into the database. We'll combine both files into a
single script that displays the form if there's no flag passed and inserts the
record into the database if another flag value is passed to the script. Just for
fun we'll even add a routine to display the information for verification before
its submitted - and use a flag for that, too. Our "flags" and values
could be just about anything you'd like, but to keep things very simple, we'll
use "Flag" and make the values numeric.</font>
<p><a href="http://www.web-savant.com/users/kathi/asp/samples/database/sample5.asp" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/sample5code.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a>      <a href="#top"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</a></font>
<p><a name="paging"><b><font face="Verdana" size="2">Paging Through Recordsets</font></b></a><font face="Verdana" size="2"><br>
Now that we've gone throught the basics of database access, let's go a step
further. The sample script above for viewing a database is fine if you only have
a few records. Imagine, however, that you have several hundred - or even
thousands - of records. Clearly the above script won't work. We will need to add
some sort of paging mechanism so we can navigate through our recordset. The key
to this is built-in features of ADO: .RecordCount, .PageSize, and
.AbsoluteCount. We'll use these features to figure out how many records are in
the recordset, specify how many records you want per page, and which page of the
recordset we want to work with. This is also an example of a script that submits
to itself; in this case the script acts upon the page number that's passed to
it. If there is no page number passed to the script, the script sets the page
number to 1. The sample script is part of the <a href="http://www.web-savant.com/users/kathi/asp/guestbook/guestbook.zip">Free
Guestbook</a> download.</font>
<p><a href="http://www.web-savant.com/users/kathi/guestbook/default.asp" target="_new"><font face="Verdana" size="2"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Working Sample</font></a><font face="Verdana" size="2">
     <a href="http://www.planet-source-code.com/vb/tutorial/asp/samples/gbpaging.html" target="_new"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">View
Code</a>      <a href="#top"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</a></font>
<p><a name="summary"><b><font face="Verdana" size="2">Summary</font></b></a><font face="Verdana" size="2"><br>
In this section, we've learned how to connect to a database, perform a simple
query, and display the contents on a dynamically generated web page. We've also
learned how to modify records in the database, and how to delete records. Now
that we've got the basic skills to read and write to a database, its not much of
a jump to extend these skills to writing a database-generated guestbook. The
next page will show you how...</font>
<p><a href="#top"><font face="Verdana" size="2"><img border="0" height="10" src="http://www.planet-source-code.com/vb/tutorial/asp/images/top.gif" width="10">  Top
of Page</font></a><font face="Verdana" size="2">       <a href="http://www.planet-source-code.com/vb/tutorial/asp/default.asp?txtTutorialName=DatabaseBasics2.asp"><img border="0" src="http://www.planet-source-code.com/vb/tutorial/asp/images/compact.gif" width="11" height="11">Building
a Guestbook</a></font></p>

