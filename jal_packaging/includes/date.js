function chBrowser(){   // Check browser NC or IE
  var bt="IE"
  if (navigator.appName=="Netscape")
 bt = "NC";
  return bt;}
var StartYear=1900;  // start year of Javascriptvar
 today=new Date();
var Jour=today.getDay();
var JourMois=today.getDate();
var Mois=today.getMonth();
var Annee=today.getYear();
if(chBrowser() == "NC")
	Annee += StartYear;
var JourLettres=new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
var MoisLettres=new Array("January","February","March","April","May","June","July","August","September","October","November","December");
document.write(JourLettres[Jour]+", "+MoisLettres[Mois]+" "+JourMois+", "+Annee);