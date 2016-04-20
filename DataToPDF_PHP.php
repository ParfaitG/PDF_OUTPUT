<?php

// SET CURRENT DIRECTORY
$cd = dirname(__FILE__);

// READ CSV FILE
$handle = fopen($cd."/DATA/EnergyConsumptionBySector1949-2015.csv", "r");    

$row = 0; $energydata = [];

while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
    $row++;
    if($row == 1){ $row++; continue; }
    
        $innerdata = [];
        $innerdata[] = substr(strval($data[1]), 0, 4);          // YEAR
        
        if (substr(strval($data[1]), 4, 1) == "0"){			    
            $innerdata[] = substr(strval($data[1]), 5, 1);      // MONTH 0 - 9
        } else {
            $innerdata[] = substr(strval($data[1]), 4, 2);      // MONTH 10 - 12
        }			
        $innerdata[] = $data[2];                                // VALUE
        $innerdata[] = $data[4];                                // DESCRIPTION			
      
        $energydata[] = $innerdata;
}

fclose($handle);

// TRANSPOSE DATA
$energypivot = [];
for ($yr = 1949; $yr <= 2015; $yr++){
    
    for ($mo = 1; $mo <= 13; $mo++){
        $innerdata = [];                        
        $innerdata[] = $yr;
        if ($mo == 13) {
            $innerdata[] = 'Total';
        } else {
            $innerdata[] = date("F", mktime(0, 0, 0, $mo, 15));
        }
        
        foreach($energydata as $e){
            
            if ($e[0]==$yr & $e[1]==$mo){
                                
                switch($e[3]) {			
                    case "Primary Energy Consumed by the Residential Sector": $innerdata[] = $e[2]; break;
                    case "Total Energy Consumed by the Residential Sector": $innerdata[] = $e[2]; break;		    
                    case "Primary Energy Consumed by the Commercial Sector": $innerdata[] = $e[2]; break;		    
                    case "Total Energy Consumed by the Commercial Sector": $innerdata[] = $e[2]; break;
                    case "Primary Energy Consumed by the Industrial Sector": $innerdata[] = $e[2]; break;
                    case "Total Energy Consumed by the Industrial Sector": $innerdata[] = $e[2]; break;
                    case "Primary Energy Consumed by the Transportation Sector": $innerdata[] = $e[2]; break;
                    case "Total Energy Consumed by the Transportation Sector": $innerdata[] = $e[2]; break;
                    case "Primary Energy Consumed by the Electric Power Sector": $innerdata[] = $e[2]; break;
                    case "Energy Consumption Balancing Item": $innerdata[] = $e[2]; break;
                    case "Primary Energy Consumption Total": $innerdata[] = $e[2]; break;
                }
                                
            }
            
        }
        if (sizeof($innerdata) > 2) {
            $energypivot[] = $innerdata;
        }   
    }   
}

// INITIALIZE DOM DOCUMENT
$dom = new DOMDocument('1.0', 'UTF-8');
$dom->substituteEntities = false;
$dom->formatOutput = true;
$dom->preserveWhiteSpace = false;

// CREATE ROOT
$root = $dom->createElement("html");
$root = $dom->appendChild($root);

$head = $dom->createElement("head");
$head = $root->appendChild($head);

$stylestr = "body{
                    margin:15px; padding:50px;
                    font-family:Arial, Helvetica, sans-serif; font-size:88%;
            }
            h2,h3 {
                    font:Arial black; color: #383838;
                    valign: top;
            }
            .yeartitle {
            	page-break-before: always;
            }
            img {
                  float: right;
            }
            table, tr, td, th, thead, tbody, tfoot {
                page-break-inside: avoid !important;
            }
            table{
                    width:100%; font-size:13px;
                    border-collapse:collapse;
                    text-align: right;
            }
            th{ color: #383838 ; padding:2px; text-align:right; }
            td{ padding: 2px 5px 2px 5px; }
            
            tr.headerrow{
                    border-bottom: 2px solid #A8A8A8; 
            }
            tr.even{
                    background-color: #F0F0F0;
            }
            .footer{
                    text-align: right;
                    color: #A8A8A8;
                    font-size: 12px;
                    margin-top: 10px;
            }";

$style = $dom->createElement("style", $stylestr);
$style->setAttribute("type", "text/css");
$style->setAttribute("media", "all");
$style = $head->appendChild($style);

$body = $dom->createElement("body");
$body = $root->appendChild($body);

$h1 = $dom->createElement("h1", "U.S. Energy Consumption 1949 - 2015");
$h1 = $body->appendChild($h1);

$img = $dom->createElement("img");
$img->setAttribute("src", "EnergyIcon.png");
$img->setAttribute("alt", "Energy icon");
$img = $h1->appendChild($img);

function runHeaders($doc, $table){    
    $cols = array("Year", "Month", "Residential Sector Primary", "Residential Sector Total",
                  "Commercial Sector Primary", "Commercial Sector Total", "Industrial Sector Primary", "Industrial Sector Total",
                  "Transportation Sector Primary", "Transportation Sector Total", "Electric Power Sector Primary",
                  "Energy Consumption Balancing Item", "Grand Consumption Total");    
    
    $tableheader = $doc->createElement("tr");
    $tableheader->setAttribute("class", "headerrow");
    $tableheader = $table->appendChild($tableheader);
    
    foreach($cols as $c){        
        $th = $doc->createElement("th", $c);
        $th = $tableheader->appendChild($th);
    }
    return $tableheader;
}

$earlytable =  $dom->createElement("table");
$earlytable = $body->appendChild($earlytable);
$earlyheaders = runHeaders($dom, $earlytable);

foreach($energypivot as $row){
    
    if ($row[0] <= 1972) {        
        $tr = $dom->createElement("tr");
        if ($row[0] % 2 == 0){
            $tr->setAttribute("class", "even"); 
        }        
        $tr = $earlytable->appendChild($tr);
    
        for($i=0; $i<=12; $i++){                
            $td = $dom->createElement("td", $row[$i]);
            $td = $tr->appendChild($td);
        }
    }
}
$footer = $dom->createElement("div", "Source: EIA - U.S. Department of Energy");
$footer->setAttribute("class", "footer");
$footer = $body->appendChild($footer);


for ($yr = 1973; $yr <= 2015; $yr++){
    
    $yrtitle = $dom->createElement("h2", $yr);
    $yrtitle->setAttribute("class", "yeartitle");
    $yrtitle = $body->appendChild($yrtitle);
    
    $yrimg = $dom->createElement("img");
    $yrimg->setAttribute("src", "EnergyIcon.png");
    $yrimg->setAttribute("alt", "Energy icon");
    $yrimg = $yrtitle->appendChild($yrimg);
    
    $latertable =  $dom->createElement("table");
    $latertable = $body->appendChild($latertable);
    $laterheaders = runHeaders($dom, $latertable);
    
    for($e = 0; $e < sizeof($energypivot); $e++){
        
        if ($energypivot[$e][0] == $yr) {
            $tr = $dom->createElement("tr");
            if ($e % 2 != 0){
               $tr->setAttribute("class", "even"); 
            }
            $tr = $latertable->appendChild($tr);
            
            for($i=0; $i<=12; $i++){                
                $td = $dom->createElement("td", $energypivot[$e][$i]);
                $td = $tr->appendChild($td);
            }
        }
    }
    
    $footer = $dom->createElement("div", "Source: EIA - U.S. Department of Energy");
    $footer->setAttribute("class", "footer");
    $footer = $body->appendChild($footer);
        
}

// OUTPUT HTML TO FILE
file_put_contents($cd. "/DATA/Output_PHP.html", $dom->saveHTML());

// CONVERT HTML TO PDF
exec(' wkhtmltopdf -O landscape "' . $cd.'/DATA/Output_PHP.html" "'. $cd.'/DATA/Output_PHP.pdf"');

echo "\nSuccessfully processed CSV data into PDF!\n";

?>


