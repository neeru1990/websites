<!--
var scrollerwidth=190		//width
var scrollerheight=115		//height
var scrollerbgcolor='#FFFFFF'	//background colour
var scrollerbackground=''	//set to '' if you don't wish to use a background image

//configure the below variable to change the contents of the scroller
var messages=new Array()
messages[0]="<H6>Why RightFax?</H6>Captaris has over 100,000 systems installed worldwide in more than 80% of Fortune 100 companies."
messages[1]="<H6>22nd May 2003</H6>RightFax connector for Oracle CRM now available.<BR><BR><A href='rightfax/int_oracle.htm'>Click here</A> for details."
messages[2]="<H6>Did You Know?</H6>Captaris RightFax solutions have won more awards than all competitive solutions combined."
messages[3]="<H6>Brooktrout TR1034</H6>Running successfully with 2 E1s on the same server in India. <BR><BR>Fractional E1 cards available for 8 and 16 ports. <A href='brooktrout/tr1034_ov.htm'>Read more.</A>"
messages[4]="<H6>Coming Soon!</H6>RightFax's integration with Domino R6."
messages[5]="<H6>Did You Know?</H6>May 27, 2003 marked the 160th birthday of the fax machine!"
messages[6]="<H6>Zetafax</H6>Version 8 supports Microsoft's latest software platform - Windows Server 2003."
messages[7]="<H6>Aviator Software</H6>Newest Release of Document Management Software Solutions for the Lotus Market launched."
messages[8]="<H6>Top Pick</H6>ArcMentor and IBM Sametime are a natural synergy for your organisation."


///////Do not edit pass this line///////////////////////
var ie=document.all&&navigator.userAgent.indexOf("Opera")==-1
var dom=document.getElementById&&navigator.userAgent.indexOf("Opera")==-1

if (messages.length>2)
i=2
else
i=0

function move1(whichlayer){
tlayer=eval(whichlayer)
if (tlayer.top>0&&tlayer.top<=5){
tlayer.top=0
setTimeout("move1(tlayer)",6000)
setTimeout("move2(document.main.document.second)",6000)
return
}
if (tlayer.top>=tlayer.document.height*-1){
tlayer.top-=5
setTimeout("move1(tlayer)",50)
}
else{
tlayer.top=scrollerheight
tlayer.document.write(messages[i])
tlayer.document.close()
if (i==messages.length-1)
i=0
else
i++
}
}

function move2(whichlayer){
tlayer2=eval(whichlayer)
if (tlayer2.top>0&&tlayer2.top<=5){
tlayer2.top=0
setTimeout("move2(tlayer2)",6000)
setTimeout("move1(document.main.document.first)",6000)
return
}
if (tlayer2.top>=tlayer2.document.height*-1){
tlayer2.top-=5
setTimeout("move2(tlayer2)",50)
}
else{
tlayer2.top=scrollerheight
tlayer2.document.write(messages[i])
tlayer2.document.close()
if (i==messages.length-1)
i=0
else
i++
}
}

function move3(whichdiv){
tdiv=eval(whichdiv)
if (parseInt(tdiv.style.top)>0&&parseInt(tdiv.style.top)<=5){
tdiv.style.top=0
setTimeout("move3(tdiv)",6000)
setTimeout("move4(second2_obj)",6000)
return
}
if (parseInt(tdiv.style.top)>=tdiv.offsetHeight*-1){
tdiv.style.top=parseInt(tdiv.style.top)-5
setTimeout("move3(tdiv)",50)
}
else{
tdiv.style.top=scrollerheight
tdiv.innerHTML=messages[i]
if (i==messages.length-1)
i=0
else
i++
}
}

function move4(whichdiv){
tdiv2=eval(whichdiv)
if (parseInt(tdiv2.style.top)>0&&parseInt(tdiv2.style.top)<=5){
tdiv2.style.top=0
setTimeout("move4(tdiv2)",6000)
setTimeout("move3(first2_obj)",6000)
return
}
if (parseInt(tdiv2.style.top)>=tdiv2.offsetHeight*-1){
tdiv2.style.top=parseInt(tdiv2.style.top)-5
setTimeout("move4(second2_obj)",50)
}
else{
tdiv2.style.top=scrollerheight
tdiv2.innerHTML=messages[i]
if (i==messages.length-1)
i=0
else
i++
}
}

function startscroll(){
if (ie||dom){
first2_obj=ie? first2 : document.getElementById("first2")
second2_obj=ie? second2 : document.getElementById("second2")
move3(first2_obj)
second2_obj.style.top=scrollerheight
second2_obj.style.visibility='visible'
}
else if (document.layers){
document.main.visibility='show'
move1(document.main.document.first)
document.main.document.second.top=scrollerheight+5
document.main.document.second.visibility='show'
}
}
window.onload=startscroll
-->