<div align="center">

## Using Collections in VB


</div>

### Description

Explains the basics of using collections in Visual Basic. These are a very powerful and often unused feature of VB.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-06-28 15:57:42
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Intermediate
**User Rating**    |4.4 (160 globes from 36 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD72546282000\.zip](https://github.com/Planet-Source-Code/matthew-roberts-using-collections-in-vb__1-9349/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>Using Collections</TITLE>
<META NAME="Template" CONTENT="D:\Program Files\Microsoft Office\Office\html.dot">
</HEAD>
<BODY LINK="#0000ff" VLINK="#800080">
<FONT FACE="Arial" SIZE=5><P ALIGN="CENTER">Using Collections</P>
</FONT><FONT FACE="Arial" SIZE=2><P>So you have heard of Collections and may have even used them a few times. An unassuming word&#8230;collections. It doesn&#8217;t inspire much excitement in most circles, yet there are very few single words that represent such a powerful element of programming as they do. This article will outline some of the general and specific uses of collections in Visual Basic and Access. After reading it, you will hopefully have a higher respect for this often overlooked aspect of VB.</P>
<P>Just what are collections anyway? Well, they are just what their name implies. They are a logical grouping of objects in Visual Basic. The Visual Basic object model consists of objects and collections of objects. For example, you have a "Forms" collection which contains all of the forms in the application. Each form also has an Objects collection which contains all of the objects that are contained in the form. On the Access side, there is a TableDefs collection which contains all of the tables in your database, and each of these TableDefs contains a Fields collection. As you may have guessed, the Fields collections contains all of the fields that exists in each table. </P>
<P>What does this mean to the average coder? Where is the payoff for all of this organization? You are about to find out. Using the Forms example above, consider this problem:</P>
<P>For some strange reason, your client wants you to create a function that will show all of the forms in the entire application at once. You could so something like this:</P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P>Function ShowForms</P><DIR>
<DIR>
<P>	FrmSplash.Show</P>
<P>	FrmMainMenu.Show</P>
<P>	FrmSelectUser.Show</P>
<P>	FrmOpenDocument.Show</P>
<P>	&#8230;etc&#8230;.etc&#8230;.</P></DIR>
</DIR>
<P>End Function</P>
</FONT><FONT FACE="Arial" SIZE=2><P>This can be tedious if the application has 35 or so forms. And to make matters worse, they keep adding and removing forms, so you have to keep coming back and changing this function to keep from causing a compile error "Object required" every time one changes. What a pain. You could solve this entire problem by either finding a new job, talking some sense into your client (like THAT would work!) or by using this code:</P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P>Function ShowForms</P><DIR>
<DIR>
<P>	Dim frmForm as Form</P>
<P>	For each frmForm in Forms</P><DIR>
<DIR>
<P>		FrmForm.Show</P></DIR>
</DIR>
<P>	Next frmForm</P></DIR>
</DIR>
<P>End Function</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Now the client can add, remove, and change the name of as many forms as he likes without effecting the operation of the application. By looping through (or "iterating") the collection, you have made you code immune to the whims of your client. Lets look at how this works by examining each statement. </P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P>Dim frmForm as Form</P>
</FONT><FONT FACE="Arial" SIZE=2><P>This statement creates an object variable that will hold each form object as we iterate through the Forms collection. It is basically a temporary storage space for a form object.</P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P>For each frmForm in Forms</P>
</FONT><FONT FACE="Arial" SIZE=2><P>If you haven&#8217;t yet started using the For Each &#8230; Next statement, you need to get with the program. It works just like the old Basic/VB For&#8230;Next, but it does it with objects instead of variables. This is the heart of working with collections.</P>
</FONT><FONT FACE="Arial" SIZE="2" COLOR="#000080"><P>FrmForm.Show</P>
</FONT><FONT FACE="Arial" SIZE=2><P> </P>
<P>This magic little statement takes the place of all of those other .show statements in the prior example. With each pass through the For Each&#8230;Next loop, the object variable frmForm is reassigned to contain the current form object. So when you say "frmForm.Show", VB interprets it as frmSplash.Show, frmMainMenu.Show, or whatever form is currently being proccessed. </P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P> </P>
<P>Next frmForm</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Wraps up the loop. This will return execution back to the For Each&#8230; statement above it. Code execution will pass through this loop once for each form in the Forms collection.</P>
<P> </P>
<P>Now that you understand the basic logic of iterating through collections, you can see how this could be put to practical use. To cascade all open forms on the screen, you could modify the code to this:</P>
</FONT><FONT FACE="Arial" SIZE="2" COLOR="#000080"><P>Function CascadeForms</P>
<P>Dim intTop As Integer</P>
<P>Dim intLeft As Integer</P>
<P>Dim frmForm as Form</P>
<P>For each frmForm in Forms</P>
<P>If frmForm.Visible = True Then</P><DIR>
<DIR>
<P>		IntT</FONT><FONT FACE="Arial" SIZE="2">op = intTop + 100</P>
<P>		IntLeft = IntLeft + 100</P>
<P>		FrmForm.Top = IntTop</P>
<P>		FrmForm.Left = IntLeft</FONT><FONT FACE="Arial" SIZE="2" COLOR="#000080">				</P>
<P>	End if</P></DIR>
</DIR>
<P>Next frmForm</P>
<P>End Function</P>
</FONT><FONT FACE="Arial" SIZE="2"><P>This code will place forms over each other in cascade style, starting at coordinates 100,100 and moving down and to the right in increments of 100. It took almost as many letters to explain it as it does to write it!</P>
<P>The thing to note in this example is that ALL of the forms&#8217; properties and functions are available as you loop though the collection. For example, you could have changed the caption of each one or the border style of only certain ones. </P>
<P>OK&#8230;enough about forms. Where else can these really cool collections be used? How about within a form? This code will print a list of all objects on a form to the debug window:</P>
</FONT><FONT FACE="Arial" SIZE="2" COLOR="#000080"><P>Function ShowObjects</P><DIR>
<DIR>
<P>	Dim objObject as Object</P>
<P>	For each objObject in Me</P><DIR>
<DIR>
<P>		Debug.Print objObject.Name</P></DIR>
</DIR>
<P>	Next objObject</P></DIR>
</DIR>
<P>End Function</P>
</FONT><FONT FACE="Arial" SIZE="2"><P>This will work whether you have one or 1000 objects on a form&#8230;although I wouldn&#8217;t recommend putting that many controls on a single form&#8230;but hey, it would work with it! </P>
<P>The thing to note in the above code (besides the obvious compactness of it) is the use of the Me keyword. This is important. Me translates in VB to "Whichever form this code is running in". It is used to reference the Objects collection for the current form. This means that you could copy this code from one form and paste it directly into another and it would work with NO code changes. Here is a more practical example of the objects collection:</P>
<P>You have a form with 25 text boxes on it and you want to automatically center them when the user resizes the form. You could write some pretty painful code to do this, or you could do this:</P>
</FONT><FONT FACE="Arial" SIZE="2" COLOR="#000080"><P>Private Sub Form_Resize()</P><DIR>
<DIR>
<P>	Dim objObject as Object</P>
<P>	For each objObject in Me</P><DIR>
<DIR>
<P>		ObjObject.Left = (Me.Width / 2) - (objObject.Width / 2)</P></DIR>
</DIR>
<P>	Next objObject</P></DIR>
</DIR>
<P>End sub</P>
</FONT><FONT FACE="Arial" SIZE="2"><P>This code will center any objects, no matter what their widths. With a little imagination, you can probably see how this same concept could be used to resized objects in a form as well. In fact, for the curious, I have already posted a sample project with the code to do just that in it. You can take a look at it by </FONT><A HREF="http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=9135"><FONT FACE="Arial" SIZE="2">clicking here</FONT></A><FONT FACE="Arial" SIZE="2">.</P>
<P>I hope you have found this article helpful. If you would like to have me post a follow up showing more advanced techniques for using collections, please leave some helpful comments and maybe a rating at </FONT>PlanetSourceCode</A><FONT FACE="Arial" SIZE="2">. </P>
<P>Have Fun!</P>
<P>M@</P></FONT>
PS: For information on using the Microsoft Jet Database collections, <a href="http://www.planetsourcecode.com/xq/ASP/txtCodeId.11529/lngWId.1/qx/vb/scripts/ShowCode.htm"> Click Here </a> to view my second collections tutorial.
</BODY>
</HTML>

