menunum=0;menus=new Array();_d=document;function addmenu(){menunum++;menus[menunum]=menu;}function dumpmenus(){mt="<script language=javascript>";for(a=1;a<menus.length;a++){mt+=" menu"+a+"=menus["+a+"];"}mt+="<\/script>";_d.write(mt)}
effect = "Fade(duration=0.2);Alpha(style=0,opacity=100);Shadow(color='#667799', Direction=135, Strength=0)"

timegap=500		// The time delay for menus to remain visible
followspeed=5		// Follow Scrolling speed
followrate=40		// Follow Scrolling Rate
suboffset_top=4;	// Sub menu offset Top position 
suboffset_left=0;	// Sub menu offset Left position
closeOnClick = true

style1=[		// style1 is an array of properties. You can have as many property arrays as you need. This means that menus can have their own style.
"FFFFFF",		// Mouse Off Font Color
"AAAAAA",		// Mouse Off Background Color
"FFFFFF",		// Mouse On Font Color
"666666",		// Mouse On Background Color
"EEEEEE",		// Menu Border Color 
11,			// Font Size in pixels
"normal",		// Font Style (italic or normal)
"bold",			// Font Weight (bold or normal)
"Verdana, Arial",	// Font Name
4,			// Menu Item Padding
"../images/menu_arrow.gif",// Sub Menu Image (Leave this blank if not needed)
,			// 3D Border & Separator bar
"66ffff",		// 3D High Color
"000099",		// 3D Low Color
"444444",		// Current Page Item Font Color (leave this blank to disable)
"EEEEEE",		// Current Page Item Background Color (leave this blank to disable)
"",			// Top Bar image (Leave this blank to disable)
"ffffff",		// Menu Header Font Color (Leave blank if headers are not needed)
"000099",		// Menu Header Background Color (Leave blank if headers are not needed)
"EEEEEE",		// Menu Item Separator Color
]


addmenu(menu=[		// This is the array that contains your menu properties and details
"mainmenu",		// Menu Name - This is needed in order for the menu to be called
140,			// Menu Top - The Top position of the menu in pixels
0,			// Menu Left - The Left position of the menu in pixels
125,			// Menu Width - Menus width in pixels
0,			// Menu Border Width 
"",			// Screen Position - here you can use "center;left;right;middle;top;bottom" or a combination of "center:middle"
style1,			// Properties Array - this is set higher up, as above
1,			// Always Visible - allows the menu item to be visible at all time (1=on/0=off)
"left",			// Alignment - sets the menu elements text alignment, values valid here are: left, right or center
,			// Filter - Text variable for setting transitional effects on menu activation - see above for more info
1,			// Follow Scrolling - Tells the menu item to follow the user down the screen (visible at all times) (1=on/0=off)
0, 			// Horizontal Menu - Tells the menu to become horizontal instead of top to bottom style (1=on/0=off)
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

,"RightFax", "show-menu=rightfax",,"",1
,"FaxCore", "show-menu=faxcore",,"",1
,"Zetafax", "show-menu=zetafax",,"",1
,"Aviator", "show-menu=aviator",,"",1
,"ArcMentor", "../arcmentor/overview.htm",,"",1
,"Brooktrout", "show-menu=brooktrout",,"",1
])
	
	addmenu(menu=["rightfax",,,150,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Overview", "show-menu=rf_ov",,,1
	,"Datasheet", "show-menu=rf_ds",,,1
	,"SAP Integration", "show-menu=rf_sap",,,1
	,"Integration FAQ's", "show-menu=rf_faq",,,1
	,"Online Demos", "../rightfax/demos.htm",,,1
	])
		addmenu(menu=["rf_ov",,,240,1,"",style1,,"",effect,,,,,,,,,,,,
		,"Family of Products", "../rightfax/ov_products.htm",,,1
		,"Enterprise & Network Fax", "../rightfax/ov_ent.htm",,,1
		,"Universal Information Exchange", "../rightfax/ov_uie.htm",,,1
		,"Secure E-mail Delivery", "../rightfax/ov_secureemail.htm",,,1
		,"Fax Boards", "../rightfax/ov_faxboards.htm",,,1
		,"System Requirements", "../rightfax/ov_sysreq.htm",,,1
		])

		addmenu(menu=["rf_ds",,,240,1,"",style1,,"",effect,,,,,,,,,,,,
		,"IP Plus Connector & Services", "../rightfax/ds_ipplus.htm",,,1
		,"Universal Information Exchange", "../rightfax/ds_uie.htm",,,1
		])

		addmenu(menu=["rf_sap",,,170,1,"",style1,,"",effect,,,,,,,,,,,,
		,"Overview", "../rightfax/sap_overview.htm",,,1
		,"Features", "../rightfax/sap_features.htm",,,1
		,"How It Works", "../rightfax/sap_howitworks.htm",,,1
		,"Business Integration", "../rightfax/sap_business.htm",,,1
		,"Enterprise Integration", "../rightfax/sap_enterprise.htm",,,1
		])

		addmenu(menu=["rf_faq",,,140,1,"",style1,,"",effect,,,,,,,,,,,,
		,"BAAN IV C3", "../rightfax/int_baan.htm",,,1
		,"FileNET", "../rightfax/int_filenet.htm",,,1
		,"Hewlett Packard", "../rightfax/int_hp.htm",,,1
		,"IP Plus", "../rightfax/int_ipplus.htm",,,1
		,"Lotus", "../rightfax/int_lotus.htm",,,1
		,"Microsoft", "../rightfax/int_microsoft.htm",,,1
		,"Oracle CRM", "../rightfax/int_oracle.htm",,,1
		,"Xerox", "../rightfax/int_xerox.htm",,,1
		])
	
	addmenu(menu=["faxcore",,,110,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Overview", "../faxcore/overview.htm",,,1
	,"Features", "../faxcore/features.htm",,,1
	,"Benefits", "../faxcore/benefits.htm",,,1
	,"Client", "../faxcore/client.htm",,,1
	,"Server", "../faxcore/server.htm",,,1
	,"Services", "../faxcore/services.htm",,,1
	,"Hardware", "../faxcore/hardware.htm",,,1
	,"PDF Brochure","../faxcore/axacore.pdf",,,1
	])

	addmenu(menu=["zetafax",,,110,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Overview", "../zetafax/overview.htm",,,1
	,"Benefits", "../zetafax/benefits.htm",,,1
	,"Capabilities", "../zetafax/capable.htm",,,1
	,"E-Mail Gateway", "../zetafax/gateway.htm",,,1
	])

	addmenu(menu=["aviator",,,170,1,"",style1,,"",effect,,,,,,,,,,,,
	,"Attach", "../aviator/attach.htm",,,1
	,"Team / Enterprise", "../aviator/team.htm",,,1
	,"Documents In Context", "../aviator/docs.htm",,,1
	,"FAQ's", "../aviator/faqs.htm",,,1
	,"PDF Case Study", "../aviator/dic_at_work.pdf",,,1
	])

	addmenu(menu=["brooktrout",,,110,1,"",style1,,"",effect,,,,,,,,,,,,
	,"TR1034™", "show-menu=tr1034",,,1
	,"TR114™", "show-menu=tr114",,,1
	,"TruFax®", "show-menu=trufax",,,1
	,"Whitepapers", "../brooktrout/whitepapers.htm",,,1
	])
		addmenu(menu=["tr1034",,,150,1,"",style1,,"",effect,,,,,,,,,,,,
		,"Overview", "../brooktrout/tr1034_ov.htm",,,1
		,"Features & Benefits", "../brooktrout/tr1034_featben.htm",,,1
		,"Specifications", "../brooktrout/tr1034_specs.htm",,,1
		])

		addmenu(menu=["tr114",,,240,1,"",style1,,"",effect,,,,,,,,,,,,
		,"Intelligent Fax and Voice Boards", "../brooktrout/tr114_faxvoice.htm",,,1
		,"Analog", "../brooktrout/tr114_analog.htm",,,1
		,"ISDN PRI T1/E1 & ISDN BRI", "../brooktrout/tr114_isdn.htm",,,1
		])

		addmenu(menu=["trufax",,,150,1,"",style1,,"",effect,,,,,,,,,,,,
		,"Overview", "../brooktrout/trufax_ov.htm",,,1
		,"Features & Benefits", "../brooktrout/trufax_featben.htm",,,1
		,"Specifications", "../brooktrout/trufax_specs.htm",,,1
		])
dumpmenus()