menunum=0;menus=new Array();_d=document;function addmenu(){menunum++;menus[menunum]=menu;}function dumpmenus(){mt="<script language=javascript>";for(a=1;a<menus.length;a++){mt+=" menu"+a+"=menus["+a+"];"}mt+="<\/script>";_d.write(mt)}
effect = "Fade(duration=0.2);Alpha(style=0,opacity=90);Shadow(color='#6F9F13', Direction=0, Strength=4)"

timegap=500		// The time delay for menus to remain visible
followspeed=5		// Follow Scrolling speed
followrate=40		// Follow Scrolling Rate
suboffset_top=4;	// Sub menu offset Top position 
suboffset_left=0;	// Sub menu offset Left position
closeOnClick = true

style1=[		// style1 is an array of properties. You can have as many property arrays as you need. This means that menus can have their own style.
"FFFFFF",		// Mouse Off Font Color
"4B6181",		// Mouse Off Background Color
"FFFFFF",		// Mouse On Font Color
"273445",		// Mouse On Background Color
"FFFFFF",		// Menu Border Color 
11,			// Font Size in pixels
"normal",		// Font Style (italic or normal)
"bold",			// Font Weight (bold or normal)
"Arial, Verdana",	// Font Name
3,			// Menu Item Padding
"images/arrow.gif",	// Sub Menu Image (Leave this blank if not needed)
,			// 3D Border & Separator bar
"CCCCCC",		// 3D High Color
"CCCCCC",		// 3D Low Color
"000000",		// Current Page Item Font Color (leave this blank to disable)
"D7DDE8",		// Current Page Item Background Color (leave this blank to disable)
"",			// Top Bar image (Leave this blank to disable)
"FFFFFF",		// Menu Header Font Color (Leave blank if headers are not needed)
"FFFFFF",		// Menu Header Background Color (Leave blank if headers are not needed)
"DDDDDD",		// Menu Item Separator Color
]

addmenu(menu=[		// This is the array that contains your menu properties and details
"mainmenu",		// Menu Name - This is needed in order for the menu to be called
57,			// Menu Top - The Top position of the menu in pixels
0,			// Menu Left - The Left position of the menu in pixels
120,			// Menu Width - Menus width in pixels
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

,"<img src=images/bullet.gif height=8 width=8 border=0>&nbsp;About JAL", "show-menu=about",,"",1
,"<img src=images/bullet.gif height=8 width=8 border=0>&nbsp;Products", "show-menu=products",,"",1
,"<img src=images/bullet.gif height=8 width=8 border=0>&nbsp;Contact Us", "show-menu=contact",,"",1
])

	addmenu(menu=["about",,,120,1,,style1,0,"left",effect,0,,,,,,,,,,,
	,"Home","index.htm",,"Home Page",1
	//,"Overview","overview.htm",,"Overview",1
	,"Legal Notice","legal_notice.htm",,"Legal Notice",1
	,"Privacy Policy","privacy_policy.htm",,"Privacy Policy",1		
	])

	addmenu(menu=["products",,,190,1,,style1,0,"left",effect,0,,,,,,,,,,,
	,"Aluminum Collapsible Tubes","act.htm",,"JAL Extrusion > Aluminum Collapsible Tubes",1
	,"HIPS/PVC Trays","hips_pvc.htm",,"JAL Extrusion > HIPS/PVC Trays",1
	,"Paper Rondo Trays","rondo.htm",,"JAL Ampoules > Paper Rondo Trays",1
	,"Infrastructure","infrastructure.htm",,"Infrastructure",1
	,"Work Process","work_process.htm",,"Work Process",1	
	,"Quality Control","quality_control.htm",,"Quality Control",1
	])

	addmenu(menu=["contact",,,120,1,,style1,0,"left",effect,0,,,,,,,,,,,
	,"Feedback","feedback.htm",,"Feedback",1	
	,"Quotation Request","quote.htm",,"Quotation Request",1
	])

dumpmenus()