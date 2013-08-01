<?php
//Looks into the directory and returns the files, no subdirectories
//The path to the style directory
    $dirpath = "\\\\premfs16\\sites\\Premium16\\Bombness\\data";
    $dh = opendir($dirpath);
       while (false !== ($file = readdir($dh))) {
//Don't list subdirectories
          if ((!is_dir("$dirpath/$file")) && $file != "..") {
//Truncate the file extension and capitalize the first letter
           echo htmlspecialchars($file) . '?' . filesize("$dirpath/$file") . "|";
   }
}
     closedir($dh);
?> 