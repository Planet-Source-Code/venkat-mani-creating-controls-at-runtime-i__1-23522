<div align="center">

## Creating Controls at Runtime\-I


</div>

### Description

This article will show you how to create visual basic controls like the textbox/commandbutton at runtime.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Venkat Mani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/venkat-mani.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/venkat-mani-creating-controls-at-runtime-i__1-23522/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Creating controls at runtime</title>
</head>
<body>
<p align="center"><b><font size="4">Creating controls at runtime -I</font></b></p>
<p>Many times we are faced with a situation where we want to create controls such as
TextBox , CommandButton, Label at runtime in Visual Basic . Say u wanted to create
a textBox or an array of Option Buttons but you don't know how many u might need at that point in the program&nbsp;
. Creating controls at runtime allows you the flexibility to do this and more.
You can create and&nbsp; use all the common controls that u see in your toolbar very
easily. The first step in creating a&nbsp;<br>
control at runtime involves declaring a variable which will 'hold' the control.<br>
<br>
<br>
&nbsp;1. <i><b><u>Declaration:</u></b></i><br>
<br>
It is always better to declare this variable in the <u> general declaration section</u> of a form so that&nbsp;<br>
it can be used through out the form or declare it <u> globally</u> in a module(.bas file), if u have a&nbsp;<br>
module added to your project. A good idea here is to name the variable using standard conventions.<br>
Using the&nbsp;<br>
<br>
txt prefix for a TextBox<br>
cmd prefix for a CommandButton<br>
lbl prefix for a Lable&nbsp;<br>
chk prefix for a CheckBox&nbsp;<br>
opt prefix for an OptionButton<br>
<br>
and so on. For e.g.</p>
<p><b>Dim txtInput<br>
Dim cmdInput<br>
Dim lblInput</b></p>
<p><br>
</p>
<p><br>
The Next Step involves setting the variable to a particular control like TextBox or a&nbsp;<br>
CommandButton<br>
<br>
<br>
2. <i><b><u> Preparing the variable to hold the control:</u></b></i><br>
<br>
This is the most important step while creating a control at runtime. The common format for&nbsp;<br>
creating a control is as follows<br>
<br>
<b>Set varname=frmName.Controls.Add(Control Type,Control Name,frmName)<br>
</b><br>
Here varname is the variable to which you want to set the control to ,frmName is the form name to&nbsp;<br>
which you want to add the control, Control Type is the type of control i.e "VB.TextBox" for a text&nbsp;<br>
box,"VB.CommandButton" for a command button and so on .Control Name can be the same as the&nbsp;<br>
variable name or any name.<br>
<br>
<br>
So if u wanted to create a textbox txtInput you would have to do it this way<br>
<b>Set txtInput=frmTest.Controls.Add(&quot;VB.TextBox&quot;,&quot;txtInput&quot;,frmTest)</b><br>
<br>
To create a CommandButton<br>
<b>Set cmdInput=frmTest.Controls.Add(&quot;VB.CommandButton&quot;,&quot;cmdInput&quot;,frmTest)</b><br>
<br>
To create a Label<br>
<b>Set lblInput=frmTest.Controls.Add(&quot;VB.Label&quot;,&quot;lblInput&quot;,frmTest)</b><br>
<br>
To create a CheckBox<br>
<b>Set chkInput=frmTest.Controls.Add(&quot;VB.CheckBox&quot;,&quot;chkInput&quot;,frmTest)</b><br>
<br>
To create an OptionButton<br>
<b>Set optInput=frmTest.Controls.Add(&quot;VB.OptionButton&quot;,&quot;chkInput&quot;,frmTest)</b><br>
<br>
<br>
Similarly you can add a ListBox,ComboBox,PictureBox etc<br>
<br>
<br>
<br>
3. <b><i><u>Setting the properties of the control.</u></i></b><br>
<br>
<br>
Well now that you have created the control ,you want it to be displayed,
visible.You will need to&nbsp;<br>
set a few properties before you can display the control. The 2 most important properties are the&nbsp;<br>
controlname.Left and controlname.Top properites. These 2 properties determine where your control&nbsp;<br>
will be placed on the form. It is generally a very good idea to set these properties with respect&nbsp;<br>
to the form on which they are present. For ex<br>
<br>
<b>txtInput.Left=frmTest.Left + 100</b>  Or<br>
<b>txtInput.Left=frmTest.Left/2</b><br>
<br>
and<br>
<br>
<b>txtInput.Top=frmTest.Top</b> + 100  Or<br>
<b>txtInput.Top=frmTest.Top/2</b><br>
<br>
There are 2 more properites which are equally important.They are the controlname.Width and&nbsp;<br>
controlname.Height properites.<br>
<br>
<b>txtInput.Height=25<br>
txtInput.Width=50</b><br>
In addition you may set any properties that u might need.<br>
<br>
<br>
4. <b><i><u> Displaying the control.</u></i></b><br>
<br>
This is the last step where you have got to set the .Visible property to true in order to display&nbsp;<br>
the control.<br>
<br>
Eg<br>
<b>txtInput.Visible=True<br>
cmdInput.Visible=True</b><br>
<br>
<br>
<br>
The final code should look something like this if u want to add a textbox at
runtime<br>
In the General Declaration<br>
<b>Dim txtInput</b><br>
<br>
And in any event like the form_load event or command_click event for e.g.&nbsp;
put this<br>
<br>
<b><font size="3">Set txtInput=frmTest.Controls.Add(&quot;VB.TextBox&quot;,&quot;txtInput&quot;,frmTest)<br>
txtInput.Left=frmTest.Left/2<br>
txtInput.Top=frmTest.Top/2<br>
txtInput.Height=25<br>
txtInput.Width=50<br>
txtInput.Visible=True</font></b><br>
<br>
</p>
<p>After you have created the controls you can use them as you use your controls
normally. You can set the caption, get the text inputted just as you you would
do for any control created at design time.</p>
<p>&nbsp;</p>
<p align="left"><br>
Many of you must be wondering that now We have created the controls and can set properties, but
how do we react to the events of the controls, How do we detect if the new commandbutton that was
created is clicked. This requires an advanced concept called subclassing. I shall discuss about
subclassing with respect to our controls created at runtime in my next article.<br>
<br>
<br>
If u have any comments/questions/suggestions send me a mail at <a href="mailto:venky_dude@yahoo.com">venky_dude@yahoo.com</a>
.Also check out my homepage for some cool <a href="http://www.geocities.yahoo.com/venky_dude/venkwork.htm">VB
Codes</a><br>
<br>
<br>
</p>
</body>
</html>

