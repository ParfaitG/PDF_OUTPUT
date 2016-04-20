import java.util.*;
import java.io.InputStreamReader;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.OutputKeys;

import java.time.Month;
import java.io.File;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class DataToPDF_Java {       
    
    public static void main(String[] args) {                    
            
	    String currentDir = new File("").getAbsolutePath();
	    
            ArrayList<String> msn = new ArrayList<String>();
            ArrayList<String> energydt = new ArrayList<String>();
            ArrayList<String> energyval = new ArrayList<String>();
            ArrayList<String> enduse = new ArrayList<String>();
            ArrayList<String> desc = new ArrayList<String>();                        
            ArrayList<String> unit = new ArrayList<String>();	    

	    ArrayList<ArrayList<String>> energydata = new ArrayList<ArrayList<String>>();
	    ArrayList<ArrayList<String>> energypivot = new ArrayList<ArrayList<String>>();
	    	    
	    String COLUMN_HEADER = "0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12";
            String COMMA_DELIMITER = ",";
            String NEW_LINE_SEPARATOR = "\n";
	    
            try {                
		String csvFile = currentDir + "\\DATA\\EnergyConsumptionBySector1949-2015.csv";
		BufferedReader br = null;
		String line = "";
		String cvsSplitBy = ",";
                            
		// READ IN CSV DATA
		br = new BufferedReader(new FileReader(csvFile));
		while ((line = br.readLine()) != null) {
			
			String[] csvdata = line.split(cvsSplitBy);
			
			ArrayList<String> innerdata = new ArrayList<String>();
						
			innerdata.add(csvdata[1].substring(0, 4));    	  // YEAR
			
			if (csvdata[1].substring(4, 5).equals("0")){			    
			    innerdata.add(csvdata[1].substring(5, 6));    // MONTH 0 - 9
			} else {
			    innerdata.add(csvdata[1].substring(4, 6));    // MONTH 10 - 12
			}			
			innerdata.add(csvdata[2]);                        // VALUE
			innerdata.add(csvdata[4]);                        // DESCRIPTION			
                      
			energydata.add(innerdata);
		}
		
		// PIVOT DATA
		int yr; int mo;
		String yrstr; String mostr;
		for (yr = 1949; yr <= 2015; yr++){
		    for (mo = 1; mo <= 13; mo++){			
			yrstr = Integer.toString(yr);
			mostr = Integer.toString(mo);
			ArrayList<String> edata = BuildPivot(yrstr, mostr, energydata);
			
			if (edata.size() > 2) {
			    energypivot.add(edata);
			}
		    }
		}
				
		// INITIALIZE XML DOCUMENT BUILDER
                DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();            
                DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
		Document doc = docBuilder.newDocument();
                
                // HTML ROOT
		Element rootElement = doc.createElement("html");
		doc.appendChild(rootElement);
		
		// HEAD
		Element headNode = doc.createElement("head");
                rootElement.appendChild(headNode);    
		
		// STYLE
		Element styleNode = doc.createElement("style");		
		String styleStr = String.join("\n",
				"body{", 
				"	margin:15px; padding:50px;",
				"	font-family:Arial, Helvetica, sans-serif; font-size:88%;",
				"}",
				"h2,h3 {",
				"	font:Arial black; color: #383838;",
				"	valign: top;",
				"}",
				".yeartitle{",
				"	page-break-before: always;",
				"}",
				"img {",
				"      float: right;",
				"}",
				"table, tr, td, th, thead, tbody, tfoot {",
				"	page-break-inside: avoid !important;",
				"}",
				"table{",
				"	width:100%; font-size:13px;",                  
				"	border-collapse:collapse;",
				"	text-align: right;",
				"}",
				"th{ color: #383838 ; padding:2px; text-align:right; }",
				"td{ padding: 2px 5px 2px 5px; }",
				"",
				"tr.headerrow{",
				"	border-bottom: 2px solid #A8A8A8; ",
				"}",
				"tr.even{",
				"	background-color: #F0F0F0;",                  
				"}",
				".footer{",
				"	text-align: right;",
				"	color: #A8A8A8;",
			    	"	font-size: 12px;",
				"	margin-top: 10px;",
				"}");
		
		styleNode.setAttribute("type", "text/css");
		styleNode.setAttribute("media", "all");
		styleNode.appendChild(doc.createTextNode(styleStr));
                headNode.appendChild(styleNode); 		
		
		// BODY
		Element bodyNode = doc.createElement("body");
                rootElement.appendChild(bodyNode);    
		
		// H1
		Element h1Node = doc.createElement("h1");
		h1Node.appendChild(doc.createTextNode("U.S. Energy Consumption 1949 - 2015"));
		bodyNode.appendChild(h1Node);   
		
		// IMG
		Element imgNode = doc.createElement("img");
		imgNode.setAttribute("src", "EnergyIcon.png");
		imgNode.setAttribute("alt", "Energy icon");
                h1Node.appendChild(imgNode);			
		
		// 1949 - 1972 TABLE			
		Element earlytableNode = doc.createElement("table");
		bodyNode.appendChild(earlytableNode);
		
		Element hrNode = RunHeaders(doc, earlytableNode);
		
		for (int j=0; j < energypivot.size(); j++){
		    
		    if (Integer.parseInt(energypivot.get(j).get(0)) <= 1972) {			
			
			// DATA ROWS
			Element drNode = doc.createElement("tr");
			earlytableNode.appendChild(drNode);		    
			
			if (j % 2 != 0) {
			    drNode.setAttribute("class", "even");
			}
			
			for (int e=0; e <= 12; e++){			    
			    Element dataNode = doc.createElement("td");
			    dataNode.appendChild(doc.createTextNode(energypivot.get(j).get(e)));
			    drNode.appendChild(dataNode);
			}
		    }
		}
		
		Element footer = doc.createElement("div");
		footer.setAttribute("class", "footer");
		footer.appendChild(doc.createTextNode("Source: EIA - U.S. Department of Energy"));
		bodyNode.appendChild(footer);
		
		// POST 1973 TABLES
		for (yr = 1973; yr <= 2015; yr++){
		    
		    Element yeartitle = doc.createElement("h2");
		    yeartitle.setAttribute("class", "yeartitle");
		    yeartitle.appendChild(doc.createTextNode(Integer.toString(yr)));
		    bodyNode.appendChild(yeartitle);
		    
		    Element yrimgNode = doc.createElement("img");
		    yrimgNode.setAttribute("src", "EnergyIcon.png");
		    yrimgNode.setAttribute("alt", "Energy icon");
		    yeartitle.appendChild(yrimgNode);
		    		    
		    Element latertableNode = doc.createElement("table");
		    bodyNode.appendChild(latertableNode);
		    
		    Element laterhrNode = RunHeaders(doc, latertableNode);
		    
		    for (int j=0; j < energypivot.size(); j++) {		
		    
			if (Integer.parseInt(energypivot.get(j).get(0)) == yr) {
			    
			    // DATA ROWS
			    Element laterdrNode = doc.createElement("tr");
			    latertableNode.appendChild(laterdrNode);		    
			    
			    if (j % 2 != 0) {
				laterdrNode.setAttribute("class", "even");
			    }
			    
			    for (int e=0; e <= 12; e++){				
				Element laterdataNode = doc.createElement("td");
				laterdataNode.appendChild(doc.createTextNode(energypivot.get(j).get(e)));
				laterdrNode.appendChild(laterdataNode);
			    }
			}
		    }		    
		    Element laterfooter = doc.createElement("div");
		    laterfooter.setAttribute("class", "footer");
		    laterfooter.appendChild(doc.createTextNode("Source: EIA - U.S. Department of Energy"));
		    bodyNode.appendChild(laterfooter);
		}		

                // OUTPUT CONTENT TO HTML
		TransformerFactory transformerFactory = TransformerFactory.newInstance();                
		Transformer transformer = transformerFactory.newTransformer();
                transformer.setOutputProperty(OutputKeys.INDENT, "yes");
                transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
                
		DOMSource source = new DOMSource(doc);
		StreamResult result = new StreamResult(new File(currentDir + "\\DATA\\Output_Java.html"));		
		transformer.transform(source, result);                
		
		// OUTPUT TO PDF FILE 				
    		List<String> command = new ArrayList<String>();
	        command.add("wkhtmltopdf");
		command.add("-O");
		command.add("landscape");
		command.add("\"" + currentDir + "\\DATA\\Output_Java.html" + "\"");
		command.add("\"" + currentDir + "\\DATA\\Output_Java.pdf" + "\"");
		
		ProcessBuilder pb = new ProcessBuilder(command);		
		Process p = pb.start();
				
		InputStreamReader esr = new InputStreamReader(p.getErrorStream());
		BufferedReader errStreamReader = new BufferedReader(esr);
		
		String outputline = null;
		while ((outputline = errStreamReader.readLine()) != null) {
		    System.out.println(outputline);		
		}
				
		System.out.println("Successfully processed CSV data to PDF!");
		
	    } catch (FileNotFoundException ffe) {
                System.out.println(ffe.getMessage());
	    } catch (IOException ioe) {
                System.out.println(ioe.getMessage());
            } catch (ParserConfigurationException pce) {
		System.out.println(pce.getMessage());            
            } catch (TransformerException tfe) {
		System.out.println(tfe.getMessage());                
	    } 
    }
    
    private static ArrayList<String> BuildPivot(String yr, String mo, ArrayList<ArrayList<String>> energydata) {	
	    ArrayList<String> innerdata = new ArrayList<String>();
	    innerdata.add(yr);
	    if (mo.equals("13")) {
		innerdata.add("Total");
	    } else {
		innerdata.add(Month.of(Integer.parseInt(mo)).name());
	    }
	    
	    for (int i=0; i < energydata.size(); i++) {
				
		if (energydata.get(i).get(0).equals(yr) &&
		    energydata.get(i).get(1).equals(mo)) {	

		    switch(energydata.get(i).get(3)) {			
			case "Primary Energy Consumed by the Residential Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Total Energy Consumed by the Residential Sector": innerdata.add(energydata.get(i).get(2)); break;		    
			case "Primary Energy Consumed by the Commercial Sector": innerdata.add(energydata.get(i).get(2)); break;		    
			case "Total Energy Consumed by the Commercial Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Primary Energy Consumed by the Industrial Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Total Energy Consumed by the Industrial Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Primary Energy Consumed by the Transportation Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Total Energy Consumed by the Transportation Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Primary Energy Consumed by the Electric Power Sector": innerdata.add(energydata.get(i).get(2)); break;
			case "Energy Consumption Balancing Item": innerdata.add(energydata.get(i).get(2)); break;
			case "Primary Energy Consumption Total": innerdata.add(energydata.get(i).get(2)); break;
		    }
		}		
	    }	    
	    return innerdata;	    
    }
    
    private static Element RunHeaders(Document doc, Element tableNode) {
	
	    Element hrNode = doc.createElement("tr");
	    hrNode.setAttribute("class", "headerrow");
	    tableNode.appendChild(hrNode);		
	    
	    String[] colNames = {"Year", "Month", "Residential Sector Primary", "Residential Sector Total",
	                         "Commercial Sector Primary", "Commercial Sector Total", "Industrial Sector Primary", "Industrial Sector Total",
				 "Transportation Sector Primary", "Transportation Sector Total", "Electric Power Sector Primary",
				 "Energy Consumption Balancing Item", "Grand Consumption Total"};
	    
	    for(String col: colNames){
		Element thNode = doc.createElement("th");
		thNode.appendChild(doc.createTextNode(col));
		hrNode.appendChild(thNode);	    
	    }
	    		
	    return hrNode;
    }
    
}
