<!--
var scrollerwidth=230		//width
var scrollerheight=200		//height
var scrollerbgcolor='#FFFFFF'	//background colour
var scrollerbackground='images/scroll_bg.gif'	//set to '' if you don't wish to use a background image



//////configure the below variable to change the contents of the scroller
var messages=new Array()
messages[0]="<H5>Today's Promise</H5>The LORD is good to all, and His tender mercies are over all His works.<H6>Psalm 145:9</H6>"
messages[1]="<H5>Today's Commandment</H5>Now therefore, thus says the LORD of hosts: \"Consider your ways!\"<H6>Haggai 1:5</H6>"
messages[2]="<H5>Prayer Requests</H5>There is power in prayer of togetherness. We can join with you and uplift you on your prayer need. <BR><BR><A href='#'>Click here</A> to submit your prayer requests to us."
messages[3]="<H5>Share Your Miracle!</H5>Comments from Pushkar >> <BR><BR> This is just an idea. We can ask people to share their miracle stories with us. <BR><BR>We could also publish the good one's on the website."
messages[4]="<H5>Free Gifts</H5>Comments from Pushkar >> <BR><BR> Whenever we have free bibles to give away, we can alert the visitors here."


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