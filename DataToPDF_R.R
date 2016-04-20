library(XML)

setwd('C:\\Path\\To\\Working\\Directory\\')
energydf <- read.csv('DATA\\EnergyConsumptionBySector1949-2015.csv',
                     stringsAsFactors = FALSE)

energydf$Year <- as.numeric(substr(energydf$YYYYMM, 1, 4))
energydf$MonthNum <- as.numeric(substr(energydf$YYYYMM, 5, 6))
energydf$Month <- ifelse(energydf$MonthNum == 13,
                             'Total',
                             month.name[energydf$MonthNum])

energypivotdf <- reshape(energydf[,c('Year', 'Month', 'Description', 'Value')], 
                         idvar=c('Year', 'Month'),
                         timevar=c('Description'), direction='wide')
names(energypivotdf) <- gsub('Value.', '', names(energypivotdf))

energypivotdf <- energypivotdf[c('Year', 'Month',
                               'Primary Energy Consumed by the Residential Sector', 'Total Energy Consumed by the Residential Sector',
                               'Primary Energy Consumed by the Commercial Sector', 'Total Energy Consumed by the Commercial Sector',
                               'Primary Energy Consumed by the Industrial Sector', 'Total Energy Consumed by the Industrial Sector',
                               'Primary Energy Consumed by the Transportation Sector', 'Total Energy Consumed by the Transportation Sector',
                               'Primary Energy Consumed by the Electric Power Sector', 'Energy Consumption Balancing Item',
                               'Primary Energy Consumption Total')]
rm(energydf)


# CREATE XML FILE
doc = newXMLDoc()
root = newXMLNode("html", doc = doc)

head = newXMLNode("head", parent=root)
styletext = "body{ 
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
                     margin: 5px;
             }"
style = newXMLNode("style", styletext, parent=head, 
                   attrs=c(type="text/css", media="all"))

body = newXMLNode("body", parent=root)
h1 = newXMLNode("h1", "U.S. Energy Consumption 1949 - 2015", parent=body)
img = newXMLNode("img", parent=h1, attrs=c(src="EnergyIcon.png",
                                           alt="energy icon"))

table = newXMLNode("table", parent = body)
header = newXMLNode("tr", parent=table, attrs=c(class="headerrow"))

runHeaders <- function(){
  h1 = newXMLNode("th", "Year", parent = header)
  h2 = newXMLNode("th", "Month", parent = header)
  h3 = newXMLNode("th", "Residential Sector Primary", parent = header)
  h3 = newXMLNode("th", "Residential Sector Total", parent = header)
  h4 = newXMLNode("th", "Commercial Sector Primary", parent = header)
  h5 = newXMLNode("th", "Commercial Sector Total", parent = header)
  h6 = newXMLNode("th", "Industrial Sector Primary", parent = header)
  h7 = newXMLNode("th", "Industrial Sector Total", parent = header)
  h8 = newXMLNode("th", "Transportation Sector Primary", parent = header)
  h9 = newXMLNode("th", "Transportation Sector Total", parent = header)
  h10 = newXMLNode("th", "Electric Power Sector Primary", parent = header)
  h11 = newXMLNode("th", "Energy Consumption Balancing Item", parent = header)
  h12 = newXMLNode("th", "Grand Consumption Total", parent = header)
}
runHeaders()

# WRITE XML NODES AND DATA
j = 1
for (y in unique(energypivotdf$Year)){
  if (y >= 1973){
    ytitle = newXMLNode("h2", y, parent=body, attrs=c(class="yeartitle"))
    yimg = newXMLNode("img", parent=ytitle, attrs=c(src="EnergyIcon.png", 
                                                    alt="energy Icon"))    
    
    table = newXMLNode("table", parent = body)
    header = newXMLNode("tr", parent=table, attrs=c(class="headerrow"))
    runHeaders()
  }
    
  df <- energypivotdf[energypivotdf$Year==y,]
  for (i in 1:nrow(df)){
  
    tr = newXMLNode("tr", parent = table)
    if (j %% 2 == 0) {addAttributes(tr, class="even")}
  
    c1 = newXMLNode("td", df[i, 1], parent = tr)
    c2 = newXMLNode("td", df[i, 2], parent = tr)
    c3 = newXMLNode("td", df[i, 3], parent = tr)
    c4 = newXMLNode("td", df[i, 4], parent = tr)
    c5 = newXMLNode("td", df[i, 5], parent = tr)
    c6 = newXMLNode("td", df[i, 6], parent = tr)
    c7 = newXMLNode("td", df[i, 7], parent = tr)
    c8 = newXMLNode("td", df[i, 8], parent = tr)
    c9 = newXMLNode("td", df[i, 9], parent = tr)
    c10 = newXMLNode("td", df[i, 10], parent = tr)
    c11 = newXMLNode("td", df[i, 11], parent = tr)
    c12 = newXMLNode("td", df[i, 12], parent = tr)
    c13 = newXMLNode("td", df[i, 13], parent = tr)
    j = j + 1
  }
  
  if (y >= 1972){
    footer = newXMLNode("div", "Source: EIA - U.S. Department of Energy", 
                        parent=body, attrs=c(class="footer"))
  }
}

# OUTPUT HTML 
saveXML(doc, file="DATA\\Output_R.html")

# OUTPUT PDF
system(paste0('wkhtmltopdf -O landscape "', getwd(), '\\DATA\\Output_R.html" "', 
              getwd(), '\\DATA\\Output_R.pdf"'))

print("Successfully processed CSV data to PDF!")
