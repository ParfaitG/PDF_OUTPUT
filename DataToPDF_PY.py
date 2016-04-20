import os
import pandas as pd
import calendar
import lxml.etree as et, lxml.html as html

pd.set_option('display.width', 1000)
cd = os.path.dirname(os.path.abspath(__file__))

# READ IN DATA
energydf = pd.read_csv(os.path.join(cd, 'DATA/EnergyConsumptionBySector1949-2015.csv'))

# ADD NEEDED COLUMNS
energydf['Year'] = energydf['YYYYMM'].astype(str).str[:4].astype(int)
energydf['MonthNo'] = energydf['YYYYMM'].astype(str).str[-2:].astype(int)
energydf['Month'] = energydf['MonthNo'].apply(lambda x: calendar.month_name[x] if x <= 12 else 'Total')
energydf = energydf[['YYYYMM', 'Year', 'MonthNo', 'Month', 'Value', 'EndUseSector', 'Description', 'Unit']]

# PIVOT DATA
energypivot = energydf.pivot_table(index=['Year', 'MonthNo', 'Month'],
                                   columns=['Description'],
                                   values=['Value'],
                                   aggfunc=sum).reset_index()

energypivot.columns = [i if i != 'Value' else j for i,j in zip(energypivot.columns.get_level_values(0),
                                                               energypivot.columns.get_level_values(1))]

energypivot = energypivot[['Year', 'Month',
                           'Primary Energy Consumed by the Residential Sector', 'Total Energy Consumed by the Residential Sector',
                           'Primary Energy Consumed by the Commercial Sector', 'Total Energy Consumed by the Commercial Sector',
                           'Primary Energy Consumed by the Industrial Sector', 'Total Energy Consumed by the Industrial Sector',
                           'Primary Energy Consumed by the Transportation Sector', 'Total Energy Consumed by the Transportation Sector',
                           'Primary Energy Consumed by the Electric Power Sector', 'Energy Consumption Balancing Item',
                           'Primary Energy Consumption Total']]

# BUILD HTML PAGE
root = et.Element('html')
head = et.SubElement(root, "head")
style = et.SubElement(head, "style")
style.set("type", "text/css")
style.set("media", "all")
style.text = '''\
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
                  margin: 5px;
          }
          '''

body = et.SubElement(root, "body")
h1 = et.SubElement(body, "h1")
h1.text = "U.S. Energy Consumption 1949 - 2015"        
img = et.SubElement(h1, "img")
img.set("src", "EnergyIcon.png")
img.set("alt", "energy icon")

table = et.SubElement(body, "table")
header = et.SubElement(table, "tr")
header.set("class", "headerrow")

def runHeaders():
    h1 = et.SubElement(header, "th"); h1.text = "Year" 
    h2 = et.SubElement(header, "th"); h2.text = "Month"
    h3 = et.SubElement(header, "th"); h3.text = "Residential Sector Primary"
    h4 = et.SubElement(header, "th"); h4.text = "Residential Sector Total" 
    h5 = et.SubElement(header, "th"); h5.text = "Commercial Sector Primary"
    h6 = et.SubElement(header, "th"); h6.text = "Commercial Sector Total"
    h7 = et.SubElement(header, "th"); h7.text = "Industrial Sector Primary"
    h8 = et.SubElement(header, "th"); h8.text = "Industrial Sector Total"
    h9 = et.SubElement(header, "th"); h9.text = "Transportation Sector Primary" 
    h10 = et.SubElement(header, "th"); h10.text = "Transportation Sector Total"
    h11 = et.SubElement(header, "th"); h11.text = "Electric Power Sector Primary"
    h12 = et.SubElement(header, "th"); h12.text = "Energy Consumption Balancing Item"
    h13 = et.SubElement(header, "th"); h13.text = "Grand Consumption Total"
    
runHeaders()
i = 1
for year in set(energypivot['Year'].values.tolist()):
    if year >= 1973:
        yeartitle = et.SubElement(body, "h2")
        yeartitle.text = str(year)
        yeartitle.set("class", "yeartitle")

        yrimg = et.SubElement(yeartitle, "img")
        yrimg.set("src", "EnergyIcon.png")
        yrimg.set("alt", "energy icon")
        
        table = et.SubElement(body, "table")
        header = et.SubElement(table, "tr")
        header.set("class", "headerrow")
        runHeaders()

    for row in energypivot[energypivot['Year']==year].iterrows():
        tablerow = et.SubElement(table, "tr")    
        if i % 2 == 0: tablerow.set("class", "even")
        
        c1 = et.SubElement(tablerow, "td").text = str(row[1][0])    
        c2 = et.SubElement(tablerow, "td").text = row[1][1]
        c3 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][2])    
        c4 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][3])
        c5 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][4])   
        c6 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][5])   
        c7 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][6])
        c8 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][7])
        c9 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][8])
        c10 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][9])   
        c11 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][10])
        c12 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][11])   
        c13 = et.SubElement(tablerow, "td").text = "{:.3f}".format(row[1][12])
        i += 1

    if year >= 1972:
        footer = et.SubElement(body, "div")
        footer.text = "Source: EIA - U.S. Department of Energy"
        footer.set("class", "footer")

# EXPOORT TREE TO STRING
tree_out = html.tostring(root, pretty_print=True)
   
# OUTPUT HTML PAGE
xmlfile = open(os.path.join(cd, 'DATA', 'Output_PY.html'),'wb')
xmlfile.write(tree_out)
xmlfile.close()

# OUTPUT PDF PAGE
os.system('wkhtmltopdf -O landscape "{0}" "{1}'.format(\
          os.path.join(cd, 'DATA', 'Output_PY.html'),
          os.path.join(cd, 'DATA', 'Output_PY.pdf')))
          
print("Successfully processed CSV data to PDF!")
