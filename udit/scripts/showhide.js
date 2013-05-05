function showDetails(divID)
{     
   var strShow
       strShow = divID+"Show";
   var strHide
       strHide = divID+"Hide";
       
   document.all[divID].style.display = "block";
   document.all[strShow].style.display = "none";
   document.all[strHide].style.display = "block";
}

function hideDetails(divID)
{
   var strShow
       strShow = divID+"Show";
   var strHide
       strHide = divID+"Hide";

   document.all[divID].style.display = "none";
   document.all[strShow].style.display = "block";
   document.all[strHide].style.display = "none";
}