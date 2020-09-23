<div align="center">

## You too can access all the properties/methods/events of the webbrowser \+ html object w/o 1 control\!\!


</div>

### Description



How you can access all the properties, methods, and

events of the html object libray/webbrowser control

with using any webbrowser or winsock or inet control.

Most people believe that in order to use these objects

you have to load some sort of browser object

into memory..which has always been problem for me

since i know that IE has a lot of security holes.

the secret basically lies in one method of the

HTMLobject library. it is the ".CreateDocumentFromURl"

method

It goes something like this
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Evan Toder](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/evan-toder.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/evan-toder-you-too-can-access-all-the-properties-methods-events-of-the-webbrowser-html-obj__1-58386/archive/master.zip)





### Source Code

<br>
How you can access all the properties, methods, and<br>
events of the html object libray/webbrowser control<br>
with using any webbrowser or winsock or inet control.<br>
<br>
Most people believe that in order to use these objects<br>
you have to load some sort of browser object<br>
into memory..which has always been problem for me<br>
since i know that IE has a lot of security holes.<br>
<br>
the secret basically lies in one method of the <br>
HTMLobject library. it is the ".CreateDocumentFromURl"<br>
method<br>
It goes something like this<br>
<br>
<br>
</font> <font color="#006600">'</font><br>
dim start_time as long<br>
</font> <font color="#006600">'</font><br>
</font> <font color="#006600">'</font><br>
</font> <font color="#006600">'set a reference to the </font><br>
</font> <font color="#006600">'Microsoft HTML object library</font><br>
</font> <font color="#006600">'</font><br>
</font> <font color="#006600">'========================================</font><br>
<br>
sub GetDocObject(byval sUrl as string )<br>
&nbsp; dim objHTML as new htmlDocument<br>
&nbsp; dim objDoc  as htmlDocument<br>
<br>
&nbsp;&nbsp; set objDoc=objHtml.CreateObjectFromUrl(sUrl)<br>
<br>
</font> <font color="#006600">'we now have to wait for the document to be set</font><br>
</font> <font color="#006600">'you probably want some sort of time out</font><br>
</font> <font color="#006600">'value in case of unforseen network problems</font><br>
</font> <font color="#006600">'otherwise you caught in the dreaded hanging</font><br>
</font> <font color="#006600">'loop</font><br>
<br>
&nbsp;&nbsp; start_time = getTickCount </font> <font color="#006600">'API call</font><br>
<br>
&nbsp;&nbsp; while objDoc.readystate <> "complete"<br>
&nbsp;&nbsp;&nbsp;&nbsp;  </font> <font color="#006600">'5 second timeout value</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;  if (gettickcount - start_time) > 5000 then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    exit sub<br>
&nbsp;&nbsp;&nbsp;&nbsp;  end if<br>
&nbsp;&nbsp;&nbsp;&nbsp;  doevents<br>
&nbsp;&nbsp; wend<br>
<br>
&nbsp;&nbsp; </font> <font color="#006600">'reaching this point in code means</font><br>
&nbsp;&nbsp; </font> <font color="#006600">'the document object has been set</font><br>
&nbsp;&nbsp; </font> <font color="#006600">'and you can handle it any way you</font><br>
&nbsp;&nbsp; </font> <font color="#006600">'wish accessing its links, tables..whatever</font><br>
&nbsp;&nbsp; </font> <font color="#006600">'for example</font><br>
<br>
&nbsp;&nbsp;&nbsp;  dim lcnt as long<br>
&nbsp;&nbsp;&nbsp;  dim upper as long<br>
&nbsp;&nbsp;&nbsp;  dim oLink as htmlAnchorElement<br>
<br>
&nbsp;&nbsp;&nbsp;  upper = objDoc.GetElementsByTagName("A").length -1<br>
<br>
&nbsp;&nbsp;&nbsp;  if upper > 0 then<br>
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   for lcnt = 0 to upper-1<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    set oLink = objDoc.GetElementsBytagName("A")(lcnt)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    list1.additem oLink.href <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   next lcnty<br>
&nbsp;&nbsp;&nbsp;  end if<br>
<br>
<br>
end sub<br>
  </FONT>

