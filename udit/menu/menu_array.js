menunum=0;menus=new Array();_d=document;function addmenu(){menunum++;menus[menunum]=menu;}function dumpmenus(){mt="<script language=javascript>";for(a=1;a<menus.length;a++){mt+=" menu"+a+"=menus["+a+"];"}mt+="<\/script>";_d.write(mt)}
effect = "Fade(duration=0.2);Alpha(style=0,opacity=90);Shadow(color='#667799', Direction=135, Strength=0)"

timegap=500		// The time delay for menus to remain visible
followspeed=5		// Follow Scrolling speed
followrate=40		// Follow Scrolling Rate
suboffset_top=4;	// Sub menu offset Top position 
suboffset_left=0;	// Sub menu offset Left position
closeOnClick = true

style1=[		// style1 is an array of properties. You can have as many property arrays as you need. This means that menus can have their own style.
"FFFFFF",		// Mouse Off Font Color
"667799",		// Mouse Off Background Color
"FFFFFF",		// Mouse On Font Color
"414B63",		// Mouse On Background Color
"8A96B0",		// Menu Border Color 
11,			// Font Size in pixels
"normal",		// Font Style (italic or normal)
"bold",			// Font Weight (bold or normal)
"Verdana, Arial",	// Font Name
4,			// Menu Item Padding
"menu/arrow.gif",	// Sub Menu Image (Leave this blank if not needed)
,			// 3D Border & Separator bar
"66ffff",		// 3D High Color
"000099",		// 3D Low Color
"FFFFFF",		// Current Page Item Font Color (leave this blank to disable)
"99A4BB",		// Current Page Item Background Color (leave this blank to disable)
"",			// Top Bar image (Leave this blank to disable)
"ffffff",		// Menu Header Font Color (Leave blank if headers are not needed)
"000099",		// Menu Header Background Color (Leave blank if headers are not needed)
"99A4BB",		// Menu Item Separator Color
]


addmenu(menu=[		// This is the array that contains your menu properties and details
"mainmenu",		// Menu Name - This is needed in order for the menu to be called
66,			// Menu Top - The Top position of the menu in pixels
0,			// Menu Left - The Left position of the menu in pixels
100,			// Menu Width - Menus width in pixels
0,			// Menu Border Width 
"center",		// Screen Position - here you can use "center;left;right;middle;top;bottom" or a combination of "center:middle"
style1,			// Properties Array - this is set higher up, as above
1,			// Always Visible - allows the menu item to be visible at all time (1=on/0=off)
"center",		// Alignment - sets the menu elements text alignment, values valid here are: left, right or center
,			// Filter - Text variable for setting transitional effects on menu activation - see above for more info
,			// Follow Scrolling - Tells the menu item to follow the user down the screen (visible at all times) (1=on/0=off)
1, 			// Horizontal Menu - Tells the menu to become horizontal instead of top to bottom style (1=on/0=off)
,			// Keep Alive - Keeps the menu visible until the user moves over another menu or clicks elsewhere on the page (1=on/0=off)
,			// Position of TOP sub image left:center:right
,			// Set the Overall Width of Horizontal Menu to 100% & height to the specified amount (Leave blank to disable)
,			// Right To Left - Used in Hebrew for example. (1=on/0=off)
,			// Open the Menus OnClick - leave blank for OnMouseover (1=on/0=off)
,			// ID of the div you want to hide on MouseOver (useful for hiding form elements)
,			// Background image for menu when BGColor set to transparent.
,			// Scrollable Menu
,			// Reserved for future use

// "Description Text", "URL", "Alternate URL", "Status", "Separator Bar"

,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;Home", "index.php",,"Home Page",1
,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;B.Sc. IT", "show-menu=bsc",,"Passing Standards, Syllabus, etc.",1
,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;Colleges", "show-menu=colleges",,"",1
,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;Downloads", "show-menu=downloads",,"",1
,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;Events", "show-menu=events",,"",1
,"<img src=menu/bullet.gif height=8 width=8 border=0>&nbsp;Contact Us", "show-menu=contact",,"",1
])

	addmenu(menu=["bsc",
	,,130,1,"",style1,,"left",effect,,,,,,,,,,,,
	,"General","show-menu=general",,,1
	,"Course Details","show-menu=course",,,1
	])
	
	
		addmenu(menu=["general",
		,,170,1,"",style1,,"left",effect,,,,,,,,,,,,
		,"Faculty","show-menu=faculty",,,1
		,"<img src=menu/sqr.gif border=0>&nbsp;CNN","http://www.cnn.com swapimage=menu/sqr_hi.gif",,,1
		,"<img src=menu/sqr.gif border=0>&nbsp;MSNBC","http://www.msnbc.com swapimage=menu/sqr_hi.gif",,,1
		,"<img src=menu/sqr.gif border=0>&nbsp;ABC News","http://www.abcnews.com swapimage=menu/sqr_hi.gif",,,1
		,"<img src=menu/sqr.gif border=0>&nbsp;BBC News","http://news.bbc.co.uk swapimage=menu/sqr_hi.gif",,,1
		])

			addmenu(menu=["faculty",
			,,121,1,"",style1,,"left",effect,,,,,,,,,,,,
			,"CTV News","http://www.ctvnews.com",,,1
			,"Vancouver Sun","http://www.canada.com/vancouver/vancouversun/",,,1
			])
		
		addmenu(menu=["course",
		,,130,1,"",style1,,"left",effect,,,,,,,,,,,,
		,"Passing Standards","passing.htm",,"R 4411 Passing Standards",1
		,"Syllabus","syllabus.htm",,"B.Sc. IT Syllabus",1
		,"CET","cet.htm",,"Common Entrance Test",1
		,"Slashdot","http://www.slashdot.com",,,1		
		])



	addmenu(menu=["colleges",
	,,160,1,"",style1,,"left",effect,,,,,,,,,,,,
	,"<img src=menu/sqr.gif border=0>&nbsp;Domain Names","http://www.a-q.co.uk swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;Dynamic Drive","http://www.dynamicdrive.com swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;JavaScript Kit","http://www.javascriptkit.com swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;Freewarejava.Com","http://www.freewarejava.com swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;Active-X.com","http://www.active-x.com swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;Web Monkey","http://www.webmonkey.com swapimage=menu/sqr_hi.gif",,,1
	,"<img src=menu/sqr.gif border=0>&nbsp;Jars","http://www.jars.com swapimage=menu/sqr_hi.gif",,,1
	])
	
	addmenu(menu=["downloads",,,101,1,,style1,0,"left",effect,0,,,,,,,,,,,
	,"Whitepapers","http://www.download.com",,,1
	,"Notes","http://www.tucows.com",,,1
	])
	
	addmenu(menu=["events",
	,,101,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Fiesta", "index.htm",,,1
	,"Images", "http://www.yahoo.com",,,1
	,"Upcoming", "http://altavista.com",,,1
	,"Sponsors", "http://www.excite.com",,,1
	])
	
	addmenu(menu=["contact",
	,,101,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Feedback", "index.htm",,,1
	,"Address", "http://www.yahoo.com",,,1
	,"Getting There", "http://altavista.com",,,1
	])
	

dumpmenus()