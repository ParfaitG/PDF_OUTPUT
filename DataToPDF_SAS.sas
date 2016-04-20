%let fpath = C:\Path\To\Working\Directory;

** READING IN CSV FILES;
proc import	datafile = "&fpath\DATA\EnergyConsumptionBySector1949-2015.csv"
	out = EnergyData dbms = csv replace;
run;

data EnergyData;
	set EnergyData;	
	if Substr(put(YYYYMM, 6.), 5, 2) ^= '13' then
		do;
			Year = Year(input(put(YYYYMM, 6.) || '01', yymmdd10.));
			MonthNum = Month(input(put(YYYYMM, 6.) || '01', yymmdd10.));
			Month = strip(put(input(put(YYYYMM, 6.) || '01', yymmdd10.), monname12.));
		end;
	else
	 	do;
			Year = input(Substr(put(YYYYMM, 6.), 1, 4), 4.0); 
			MonthNum = 13;
			Month = 'Total';
		end;	
run;

** STRUCTURING DATA SET;
proc sort data=EnergyData out=EnergyData;
	by Year MonthNum Month;
run;

** TRANSPOSE DATA;
proc transpose data=EnergyData out=EnergyPivot(drop=_name_ MonthNum);
	by Year MonthNum Month;			
	var Value;	
	id Description;	
run;

** EXPORT DATASET TO XML FILE;
filename out "&fpath\DATA\Output_SAS.xml";

libname out xml XMLENCODING='UTF-8';                                               

data out.EnergyPivot;
 	set Work.EnergyPivot;
run;

libname out clear;

** CONVERT RAW OUTPUT XML TO HTML WITH XSLT;
filename xslfile temp;

data _null_;
  infile cards;
  input;
  file xslfile;
  put _infile_;
cards4;
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
 <xsl:output omit-xml-declaration="no" indent="yes"/>
 <xsl:strip-space elements="*"/> 
  
  <xsl:key name="yearkey" match="ENERGYPIVOT" use="Year"/>
  <xsl:key name="monthkey" match="ENERGYPIVOT" use="Month"/>
    
  <xsl:template match="TABLE">
    <html>
     <head>
        <style type="text/css" media="all">
          body{ 
                  margin:15px; padding: 20px;
                  font-family:Arial, Helvetica, sans-serif; font-size:88%; 
          }
          h1, h2 {
                  font:Arial black; color: #383838;
                  valign: top;
          }
          .yeartitle{
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
                  margin: 10px;
          }
          </style>
     </head>
     <body>
         <h1>
            U.S. Energy Consumption 1949 - 2015<img src="EnergyIcon.png" alt="energy icon"/>
         </h1>
         <table>
             <tr class="headerrow">
                <th>Year</th>
                <th>Month</th>
                <th>Residential Sector Primary</th>
                <th>Residential Sector Total</th>
                <th>Commercial Sector Primary</th>
                <th>Commercial Sector Total</th>
                <th>Industrial Sector Primary</th>
                <th>Industrial Sector Total</th>
                <th>Transportation Sector Primary</th>
                <th>Transportation Sector Total</th>
                <th>Electric Power Sector Primary</th>
                <th>Energy Consumption Balancing Item</th>
                <th>Grand Consumption Total</th>
            </tr>    
            <xsl:apply-templates select="ENERGYPIVOT[generate-id()    
                                        = generate-id(key('yearkey', Year)[1]) and number(Year) &lt; 1973]"/>
         </table>
         <div class="footer">Source: EIA - U.S. Department of Energy</div>
           <xsl:apply-templates select="ENERGYPIVOT[generate-id()    
                                        = generate-id(key('yearkey', Year)[1]) and number(Year) &gt; 1972]"/>
     </body>
    </html>
  </xsl:template>  
 
  <xsl:template match="ENERGYPIVOT[generate-id()    
                         = generate-id(key('yearkey', Year)[1]) and number(Year) &lt; 1973]">
      <xsl:variable select="number(Year)" name="year"/>                
          <tr>
              <xsl:if test="position() mod 2 = 0">
                  <xsl:attribute name="class">even</xsl:attribute>
              </xsl:if>
              <xsl:for-each select="key('yearkey', Year)/*">  
                  <td><xsl:value-of select="."/></td>
              </xsl:for-each>
          </tr>                  
  </xsl:template> 
  
  <xsl:template match="ENERGYPIVOT[generate-id()    
                         = generate-id(key('yearkey', Year)[1]) and number(Year) &gt; 1972]">
      <xsl:variable select="number(Year)" name="year"/>
      
      <h2 class="yeartitle"><xsl:value-of select="$year"/><img src="EnergyIcon.png" alt="energy icon"/></h2>      
      <table>
           <tr class="headerrow">
              <th>Year</th>
              <th>Month</th>
              <th>Residential Sector Primary</th>
              <th>Residential Sector Total</th>
              <th>Commercial Sector Primary</th>
              <th>Commercial Sector Total</th>
              <th>Industrial Sector Primary</th>
              <th>Industrial Sector Total</th>
              <th>Transportation Sector Primary</th>
              <th>Transportation Sector Total</th>
              <th>Electric Power Sector Primary</th>
              <th>Energy Consumption Balancing Item</th>
              <th>Grand Consumption Total</th>
          </tr>          
            
           <xsl:for-each select="key('yearkey', Year)/Month">            
            <tr>
              <xsl:if test="position() mod 2 = 0">
                  <xsl:attribute name="class">even</xsl:attribute>
              </xsl:if>
              
              <td><xsl:value-of select="../*[1]"/></td>
              <td><xsl:value-of select="../*[2]"/></td>
              <td><xsl:value-of select="../*[3]"/></td>
              <td><xsl:value-of select="../*[4]"/></td>
              <td><xsl:value-of select="../*[5]"/></td>
              <td><xsl:value-of select="../*[6]"/></td>
              <td><xsl:value-of select="../*[7]"/></td>
              <td><xsl:value-of select="../*[8]"/></td>
              <td><xsl:value-of select="../*[9]"/></td>
              <td><xsl:value-of select="../*[10]"/></td>
              <td><xsl:value-of select="../*[11]"/></td>
              <td><xsl:value-of select="../*[12]"/></td>
              <td><xsl:value-of select="../*[13]"/></td>
             </tr>
           </xsl:for-each>     
      </table>
      <div class="footer">Source: EIA - U.S. Department of Energy</div>      
  </xsl:template>
  
</xsl:stylesheet>  
;;;;

proc xsl 
	in="&fpath\DATA\Output_SAS.xml"
	out="&fpath\DATA\Output_SAS.html"
	xsl=xslfile;
run;

** CONVERT HTML TO PDF USING WKHTMLTOPDF;
%let in =  "&fpath\DATA\Output_SAS.html"; 
%let out =  "&fpath\DATA\Output_SAS.pdf"; 

options xwait; 
x wkhtmltopdf.exe -O landscape &in  &out; 
run; 

