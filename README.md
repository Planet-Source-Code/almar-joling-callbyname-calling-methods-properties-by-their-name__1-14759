<div align="center">

## CallByName \- Calling Methods/Properties by their name

<img src="PIC2001127725275032.jpg">
</div>

### Description

CallByName can be used to call poperties/methods by their name (string).

This article will learn you how to correctly use the "CallByName" function, using two example programs. It will show how you can Get and Let properties, and how to call methods.

Great for scripting languages!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-01-27 13:24:58
**By**             |[Almar Joling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/almar-joling.md)
**Level**          |Intermediate
**User Rating**    |4.9 (103 globes from 21 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD142191272001\.zip](https://github.com/Planet-Source-Code/almar-joling-callbyname-calling-methods-properties-by-their-name__1-14759/archive/master.zip)





### Source Code

Introduction:
When I was browsing PSC some time ago, I found a scripting language.
Somebody commented on it that it was using many "If...Then...End if" structures to process all the commands.
Well there's an alternative to this in VB6 (only!), and it will give you less work, and increases your programs overall speed!
This alternative is called... <B>CallByName</B><P>
<B>CallByName</B> allows programmers to call any function, sub or property by the name of it.
I'm going to demostrate this with a small, and easy understandable math program. This math program will be very simple, but it will effectively show you how to use <B>CallByName</B>. I've attached the source code as a zip file, so you can easily see how to use <B>CallByName</B> without having to build the enitre sample app from this tutorial!
This tutorial shows how you can call a method(Function/Sub) and how to change a property.
<H2>Tutorial 1 - Methods</H2>
Okay let's get started. In this program we are going to use the following controls:<P>
 2 Text boxes: Named txtValue1 and txtValue2<BR>
 1 Combo box: Named cmbAction<BR>
 1 Label: Named lblResult<BR>
 1 command button: Named cmdExecute<P>
I've named the form "frmMain". This is standard for all my projects. it isn't really necassary for this project though.
The placement of the controls does not really matter.
Okay double click on the form so that you get the code window, and add the following:<P>
<code><font color="#000084">Public </font><font color="black">Sub Form_Load()</font><BR>
<font color="#000084">
  With cmbAction</font><font color="black"><BR>
      .AddItem "Multiply"<BR>
      .AddItem "Minus"<BR>
      .AddItem "DivideBy"<BR>
      .AddItem "Plus"<BR>
      .ListIndex = 0<BR>
 </font>
   <font color="#000084">End With<BR>
End Sub</font></code><P>
The items added to the combobox are the "mathematical actions" we're going to use in our sample application.
<CODE>Listindex = 0</CODE> only sets the first item as an active item.<P>
Now, we have to make our command button "cmdExecute" do something. So, add the following code:<P>
<CODE>
<font color="#000084">Private Sub</font><font color="black">cmdExecute_Click()</font><BR>
  <font color="black">lblResult.Caption = CallByName(frmMain, cmbAction.Text, VbMethod, txtValue1, txtValue2)</font><BR>
<font color="#000084">End Sub</font><P>
</CODE>
This is the important part, especially for this tutorial. That's why I'm going to explain it very detailed.<BR>
The syntax of <B>CallByName</B> is as following:<P>
<DL>
<DT><CODE>Function CallByName(Object As Object, ProcName As String, CallType As VbCallType, Args() As Variant)<P></CODE></DT>
<DD>
 <CODE>Object as Object</CODE>: This is the object that contains the property/procedure you're calling by name.<BR>
	 So if you want to use the "Left" property of a command button, the object should be the command button.<BR>
	 If it's a procedure in a form, you need to put the form name here.<P>
 <CODE>ProcName As String</CODE>: When calling <B>CallByName</B> you have to specify the property/procedure you're going 	to call or modify, in this sub.<BR> So if you want to call the "Left" property of a command button, you need to put 	"Left" here.<P>
 <CODE>CallType as VbCallType</CODE>: Specify's the type of thing you're calling. A Property(VbLet,VbGet,VbSet) or a 	procedure(VbMethod).<BR> In this example we are going to use VbMethod, because we are going to call functions.<P>
 <CODE>Args() As Variant</CODE>: This is not a real array, like you might think. You just have to put all the values you 	want to use after each other (with "," as separator). They have to match the Method/Property you're calling!<BR> In our example we are going to use functions which need to values. txtValue1 and txtValue2.<BR> Now if you're going to change a property "Left" of a command button, you just specify one new value, which is going to be the new "Left" value.
</DD>
</DL>
<P>
I hope you all understand this. It looks complex the first time, but with some code, you're going to find this very easy!<BR>
We're now going to put our mathematical code into the program.<BR> It's very simple math.<BR>
I'm not all too good in Math, but the main point is that you understand how to use <B>CallByName</B><P>
Add the following code:<BR>
<font color="#000084">Public Function </font>Multiply(lngValue1 <font color="#000084">As Long</font>, lngValue2 <font color="#000084">As Long</font>)<font color="#000084"> As Long</font><BR>
   Multiply = lngValue1 * lngValue2<BR>
<font color="#000084">End Function</font><P>
<font color="#000084">Public Function </font>Minus(lngValue1 <font color="#000084">As Long</font>, lngValue2 <font color="#000084">As Long</font>)<font color="#000084"> As Long</font><BR>
   Minus = lngValue1 - lngValue2<BR>
<font color="#000084">End Function</font><P>
<font color="#000084">Public Function</font> DivideBy(lngValue1 <font color="#000084">As Long</font>, lngValue2 <font color="#000084">As Long</font>)<font color="#000084"> As Long</font><BR>
   DivideBy = lngValue1 / lngValue2<BR>
<font color="#000084">End Function</font><P>
<font color="#000084">Public Function</font> Plus(lngValue1 <font color="#000084">As Long</font>, lngValue2 <font color="#000084">As Long</font>)<font color="#000084">As Long</font><BR>
   Plus = lngValue1 + lngValue2<BR>
<font color="#000084">End Function</font><P>
Well, those functions should be self explaining. They require two values, and then they do the action represented by the Function's name.<BR>
Got everything ready? Okay run the program [F5].<BR>
Enter a number in both textboxes. Very high numbers will probably cause an "Overflow", so don't enter malicious numbers :o)
<P>
Now, when you press the command button the action you have chosen in the combo box will be executed!<BR>
Only by using the name of the Procedure, and the <B>CallByName</B> method.<P>
<H2>Tutorial 2 - Properties</H2>
<B>CallByName</B> can also be used for setting and retrieving properties. I'll show you how you do that. The source code is also available in the zipfile I earlier mentioned.<P>
The sample application will change the caption of the form (Let), enable/disable a timer (Get/Let), and move a command button around the form.<BR>
In this tutorial, we need the following controls:<BR>
1 Form: Named FrmMain. ScaleMode = VbPixel (3)!<BR>
2 Command buttons: Named cmdChangeCaption and cmdEnableTimer<BR>
1 Timer: Named tmrMove. Interval = 100<BR>
Placement does not really matter.<P>
We are going to change the Form's caption first. Add the following code to the command button named "cmdChangeCaption":<P>
<CODE><font color="#000084">Private Sub</FONT><font color="black"> cmdChangeCaption_Click()<BR>
   CallByName frmMain, "Caption", VbLet, "CallByName - Tutorial 2"<BR>
<font color="#000084">End Sub</FONT></CODE><P>
So what does this code do? Well, when you click on the command button, it will change the caption of the form to "CallByName - Tutorial 2".<BR> VbLet means that you set the property of an object.<P>
Now where are going to add some code that might look complex, but in fact it really isn't.<BR>
Add the following code to cmdEnableTimer:<P>
<CODE><font color="#000084">Private Sub</Font> cmdEnableTimer_Click()<BR>
   CallByName tmrMove, "Enabled", VbLet, <font color="#000084">Not</Font> CallByName(tmrMove, "Enabled", VbGet)<BR>
<font color="#000084">End Sub</Font></CODE><P>
This code sets the property "Enabled" of the timer. The code is made very efficient, because when you press again it will set the propery to the inverse of the current state.<BR> True-False-True-False, and so on... It retrieves the property using <B>CallByName</B> using "VbGet".<P>
At the moment, the timer does nothing. So let's change that. Add the following code to "tmrMove":<P>
<CODE><font color="#000084">Private Sub</FONT> tmrMove_Timer()<BR>
   CallByName cmdEnableTimer, "Left", VbLet, <font color="#000084">CInt(</FONT>Rnd(frmMain.ScaleWidth)<font color="#000084">)</FONT> * 100<BR>
   CallByName cmdEnableTimer, "Top", VbLet, <font color="#000084">CInt(</FONT>Rnd(frmMain.ScaleHeight)<font color="#000084">)</FONT> * 100<BR>
<font color="#000084">End Sub</FONT></CODE><P>
This code will put the command button on random places (in your form), after you press "Enable Timer".<BR>
If you click again on the button. (Or press enter when it has the focus) the Timer will disable.<BR> Pressing it again will enable it, and so forth...<P>
I hope you enjoyed my first tutorial! If there are any comments please do not hesitate to write them down!<P>
Cheers,<BR>
Almar Joling<BR>
<A HREF="mailto:ajoling@quadrantwars.com">ajoling@quadrantwars.com</A><BR>
<A HREF="http://www.quadrantwars.com">http://www.quadrantwars.com</A><BR>
(Completed on 27/01/2001)

