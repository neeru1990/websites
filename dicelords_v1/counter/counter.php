<? 
if (!$page)
{
 	$page = "index.php";
}
if (file_exists ($user.".txt"))
{
  $fa = @fopen($user.".txt", "a") or die("Could not open the file for writing!");
	$today = date( "m-d-Y, h:iA", time() );
  $data = $today." | ".$_SERVER["REMOTE_ADDR"]."\n";
	fwrite($fa, $data);
  fclose($fa);
}

header("Expires: Mon, 26 Jul 1997 05:00:00 GMT");  
header ("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT"); 
header ("Cache-Control: no-cache, must-revalidate"); // HTTP/1.1 
header ("Pragma: no-cache"); // HTTP/1.0 
header ("Content-type: image/png"); 
$im = @ImageCreate (30, 10) 
or die ("Cannot Initialize new GD image stream"); 
$white = ImageColorAllocate ($im, 255, 255, 255); 
$trans = imagecolortransparent($im,$white); 
ImagePng ($im); 
?> 