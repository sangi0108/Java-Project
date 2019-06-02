package Lib.Util;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringReader;

import java.net.URL;
import java.nio.charset.Charset;
import java.text.MessageFormat;
import java.util.ArrayList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.xml.sax.InputSource;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;
import org.json.XML;
import org.w3c.dom.CharacterData;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Automation_CMBS 
{
	static String url = "https://finsight.com/api/deals/recently-priced/?page= {0}&sector_primary=CMBS&sector_secondary=Conduit&sectors_excluded=";
	static ArrayList<Entity_CMBS> entityList;
  
	public static void main(String[] args)
	{
		try
		{
			// get the url content and populate the entity array
			entityList = new ArrayList<Entity_CMBS>();
			int i = 1;
			int pages = 0;
			while(i > 0 )
			{
			    String formattedURL = MessageFormat.format(url,i); 
				i++;
				JSONObject json = readJsonFromUrl(formattedURL);
				String xml = XML.toString(json);
				if(xml != null && !xml.isEmpty())
				{
					// make the xml well formed
					xml = "<root>" + xml + "</root>";
					Boolean result = parseJson(xml);
					if(!result)
					{
						i = 0;
					}
					else
					{
						pages++;
					}
				}
			}
			if(pages > 0)
			{
			  // write message	
			  writeToExcel();
			}
		}
		catch(Exception e)
		{
			// print error
		}		
	}
	
	public static JSONObject readJsonFromUrl(String url) throws IOException, JSONException 
	{
	    InputStream is = new URL(url).openStream();
	    try 
	    {
	      BufferedReader rd = new BufferedReader(new InputStreamReader(is, Charset.forName("UTF-8")));
	      StringBuilder sb = new StringBuilder();
	        String line = null;
	        while ((line = rd.readLine()) != null) 
	        {
	            sb.append(line + "\n");	            
	        }
	      String jsonText = sb.toString();
	      JSONObject json = new JSONObject(jsonText);
	      return json;
	    } 
	    finally 
	    {
	      is.close();
	    }		  
	}
		  
	public static void writeToExcel() 
	{
		int rownum = 0;
		try
		{
		    //Blank workbook
		    XSSFWorkbook workbook = new XSSFWorkbook();
	
		    //Create a blank sheet
		    XSSFSheet sheet = workbook.createSheet("Conduit");
		    //This data needs to be written (Object[])
		    Object[] headers = new Object[]{"Deal Name", "Sponsor Name", "Structuring Lead" , "Joint Lead" , "Issue Date" , "Class" , "$(M)" , "WAL" , "MO", "SP" , "FI", "DR" , "KR", "MS" , "FX/FL", "BNCH" , "GDNC", "SPRD" , "CPN", "YLD" , "Price"};
		    XSSFCellStyle style  = workbook.createCellStyle();
		    
		    // Cell Borders
		    IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
		    XSSFColor template = new XSSFColor(IndexedColors.BLACK , colorMap);
		    style.setBorderBottom(BorderStyle.THIN);
		    style.setBottomBorderColor(template);
		    style.setBorderLeft(BorderStyle.THIN);
		    style.setLeftBorderColor(template);
	        style.setBorderRight(BorderStyle.THIN);
	        style.setRightBorderColor(template);
	        style.setBorderTop(BorderStyle.THIN);
	        style.setTopBorderColor(template);
		    //Iterate over data and write to sheet
		    //create a row of excelsheet
		    XSSFRow row = sheet.createRow(rownum++);
	    	//get object array of particular key
	        int cellnum = 0;
	        XSSFCell cell;
	        // Print headers 
	        for (Object obj : headers) 
	        {
	        	cell = row.createCell(cellnum++);
	        	cell.setCellValue((String) obj);
	        	cell.setCellStyle(style);
	        }
	        // print the data
	        
		    for (Entity_CMBS key : entityList) 
		    {	    
		    	ArrayList<CollectionEntity_CMBS> values = key.getTableValues();	               
		        for (CollectionEntity_CMBS item : values)
		        { 		        	
		        	row = sheet.createRow(rownum++);
		        	cellnum = 0;
		        	// Deal Name
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(key.getDealName());
	 	        	
	 	            // Sponsor Name
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(key.getSponsorname());
	 	        	
	 	            // ST Lead
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(key.getStructuringLead());
	 	        	
	 	            // JT Lead
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(key.getJointLead());
	 	        	
	 	            // issue date
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(key.getIssueDate());
	 	        	
	 	        	// class
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getClasse());
	 	        	
	 	            // $(M)
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getM());
	 	        	
	 	            // WAL
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getWal());
	 	        	
	 	            // Mo
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getMo());
	 	        	
	 	            // SP
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getSP());
	 	        	
	 	            // FI
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getFI());
	 	        	
	 	            // DR
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getDR());
	 	        	
	 	            // KR
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getKR());
	 	        	
	 	            // MS
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getMS());
	 	        	
	 	            // FX/FL
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getFXFL());
	 	        	
	 	            // BNCH
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getBNCH());
	 	        	
	 	            // GDNC
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getGDNC());
	 	        	
	 	        	// SPRD
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getSPRD());
	 	        	
	 	            // CPN
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getCPN());
	 	        	
	 	            // YLD
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getYLD());
	 	        	
	 	            // Price
		        	cell = row.createCell(cellnum++);
	 	        	cell.setCellValue(item.getPrice());        	
	 	        	
		        }	        
		    }
		    
		    for(int i = 0 ; i <= 20 ; i++)
		    {
		    	sheet.autoSizeColumn(i);
		    }
	   
	        //Write the workbook in file system
	        FileOutputStream out = new FileOutputStream(new File("FinSight_Conduit.xlsx"));
	        workbook.write(out);
	        out.close();
	        System.out.println("File written successfully : " + System.getProperty("user.dir"));
	        workbook.close();
	    } 
	    catch (Exception e)
	    {
	        int x = rownum;
	        System.out.println(x);
	    }
	    
	}
	
	public static Boolean parseJson(String xmlString) 
	{	     
	    try 
	    {	       
	    	xmlString = xmlString.replace("<Pricing Speed>","<Pricing_Speed>").replace("</Pricing Speed>", "</Pricing_Speed>");
	        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
	        DocumentBuilder db = dbf.newDocumentBuilder();
	        InputSource is = new InputSource();
	        is.setCharacterStream(new StringReader(xmlString));
//	        PrintWriter writer1 = new PrintWriter(new File("C:\\installers\\testout.txt"));  
//	        writer1.write(xmlString);                                                   
//	        writer1.flush();  
//	        writer1.close();  
	        
            System.out.println(xmlString);
	        Document doc = db.parse(is);
	        NodeList nodes = doc.getElementsByTagName("results");
	        int nodelength = nodes.getLength();
            if(nodelength == 0)
            {
            	return false;
            }
            NodeList parentnode = doc.getElementsByTagName("root");
            Element parentElement = (Element)parentnode.item(0);
            NodeList parent = parentElement.getChildNodes();
	        // iterate the employees
	        for (int i = 0; i < parent.getLength(); i++) 
	        {
	           Entity_CMBS obj = new Entity_CMBS();
	           if(parent != null)
	           {
	        	   Element element = (Element)parent.item(i);
	               if(element != null)
	               {
	            	   if(element.getNodeName().equalsIgnoreCase("count"))
	            	   {
	            		   continue;
	            	   }
	            	   // Deal Name
	            	   NodeList series_name = element.getElementsByTagName("series_name");
	            	   String seriesName = "";                             
	                   Element issuer = (Element)element.getChildNodes().item(8); // issuer index : 8
	                   if(issuer != null)
	                   {
	                	   Element ticker = (Element)issuer.getChildNodes().item(1); // issuer index : 8
	                	   seriesName =  getCharacterDataFromElement(ticker);
	                   }
	                   Element beta = (Element)series_name.item(0);
	                   seriesName += " " +  getCharacterDataFromElement(beta);
	                   obj.setDealName(seriesName);
	                   
	                   // Sponsor 
	                   if(issuer != null)
	                   {
	                	   Element sponsor = (Element)issuer.getChildNodes().item(0); // sponsor index : 0
	                	   if(sponsor != null)
	                	   {
	                		   Element company = (Element)sponsor.getChildNodes().item(3);
	                		   if(company != null)
	                		   {
	                			   Element companyName = (Element)company.getChildNodes().item(3);
	                			   String cName =  getCharacterDataFromElement(companyName);
	                			   obj.setSponsorname(cName);
	                		   }
	                	   }
	                   }
	                   
	                   // Structuring Lead
	                   NodeList childs = element.getChildNodes();
	                   String st_leads = "";
	                   String jt_leads = "";
	                   ArrayList<CollectionEntity_CMBS> coll = new ArrayList<CollectionEntity_CMBS>();
	                   for(int index = 0 ; index < childs.getLength() ; index++)
	                   {
	                	   Element node = (Element)childs.item(index);                	   
	                	   if(node.getNodeName() == "structuring_leads")
	                	   {
	                		   Element abbr = (Element)node.getChildNodes().item(7);
	                		   String var =  getCharacterDataFromElement(abbr);
	                		   st_leads += var + " , ";
	                	   }
	                	   if(node.getNodeName() == "joint_leads")
	                	   {
	                		   Element abbr = (Element)node.getChildNodes().item(7);
	                		   String var =  getCharacterDataFromElement(abbr);
	                		   jt_leads += var + " , ";
	                	   }
	                	   if(node.getNodeName() == "pricing_date")
	                	   {
	                		   String var =  getCharacterDataFromElement(node);
	                		   obj.setIssueDate(var);
	                	   }
	                	   
	                	   // Table values 
	                	   if(node.getNodeName() == "tranches")
	                	   {
	                		   CollectionEntity_CMBS row = new CollectionEntity_CMBS();
	                		   NodeList childofNode = node.getChildNodes();
	                		   for(int childIndex = 0 ; childIndex < childofNode.getLength() ; childIndex++)
	                		   {
	                			   Element childNode = (Element)childofNode.item(childIndex);                	   
	                        	   if(childNode.getNodeName() == "class")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setClasse((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");                        		   
	                        	   }
	                        	   if(childNode.getNodeName() == "size")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setM((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "weighted_average_life")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setWal((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_moodys")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setMo((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_s_and_p")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setSP((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_fitch")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setFI((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_dbrs")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setDR((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_kroll")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setKR((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "rating_morningstar")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setMS((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "coupon_type")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setFXFL((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "benchmark")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setBNCH((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "guidance")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setGDNC((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "spread")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setSPRD((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "coupon")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setCPN((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "yield")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setYLD((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	   if(childNode.getNodeName() == "issue_price")
	                        	   {
	                        		   String var =  getCharacterDataFromElement(childNode);
	                        		   row.setPrice((var != null && !var.isEmpty() && !var.equalsIgnoreCase("null")) ? var : "-");
	                        	   }
	                        	  
	                		   }
	                		   coll.add(row);
	                	   }               	  
	                   }        
	                   obj.setTableValues(coll);
	            	   if(!st_leads.isEmpty())
	            	   {
	            		   st_leads = st_leads.replaceAll(", $", "");
	            		   obj.setStructuringLead(st_leads);
	            	   }
	            	   else
	            	   {
	            		   obj.setStructuringLead("-");
	            	   }
	            	   if(!jt_leads.isEmpty())
	            	   {
	            		   jt_leads = jt_leads.replaceAll(", $", "");
	            		   obj.setJointLead(jt_leads);
	            	   }
	            	   else
	            	   {
	            		   obj.setJointLead("-");
	            	   }
	            	   entityList.add(obj);	            	   
	               } 
	           }                    
	        }
	    }
	    catch (Exception e) 
	    {
	        return false;
	    }
	    
	    return true;
	      
	  }

	  public static String getCharacterDataFromElement(Element e) 
	  {
	    Node child = e.getFirstChild();
	    if (child instanceof CharacterData) 
	    {
	       CharacterData cd = (CharacterData) child;
	       return cd.getData();
	    }
	    return "-";
	  }

}

class Entity_CMBS
{
	
	private String DealName;
	private String sponsorname;
	private String StructuringLead;
	private String JointLead;
	private String IssueDate;
	private ArrayList<CollectionEntity_CMBS> TableValues;
	
	public String getDealName() {
		return DealName;
	}
	public void setDealName(String dealName) {
		DealName = dealName;
	}
	public String getSponsorname() {
		return sponsorname;
	}
	public void setSponsorname(String sponsorname) {
		this.sponsorname = sponsorname;
	}
	public String getStructuringLead() {
		return StructuringLead;
	}
	public void setStructuringLead(String structuringLead) {
		StructuringLead = structuringLead;
	}
	public String getJointLead() {
		return JointLead;
	}
	public void setJointLead(String jointLead) {
		JointLead = jointLead;
	}
	public String getIssueDate() {
		return IssueDate;
	}
	public void setIssueDate(String issueDate) {
		IssueDate = issueDate;
	}
	public ArrayList<CollectionEntity_CMBS> getTableValues() {
		return TableValues;
	}
	public void setTableValues(ArrayList<CollectionEntity_CMBS> tableValues) {
		TableValues = tableValues;
	}
	
}

class CollectionEntity_CMBS
{
	private String Classe;
	private String M;
	private String Wal;
	private String Mo;
	private String SP;
	private String FI;
	private String DR;
	private String KR;
	private String MS;
	private String FXFL;
	private String BNCH;
	private String GDNC;
	private String SPRD;
	private String CPN;
	private String YLD;
	private String Price;
	
	public String getClasse() {
		return Classe;
	}
	public void setClasse(String classe) {
		Classe = classe;
	}
	public String getM() {
		return M;
	}
	public void setM(String m) {
		M = m;
	}
	public String getWal() {
		return Wal;
	}
	public void setWal(String wal) {
		Wal = wal;
	}
	public String getMo() {
		return Mo;
	}
	public void setMo(String mo) {
		Mo = mo;
	}
	public String getSP() {
		return SP;
	}
	public void setSP(String sP) {
		SP = sP;
	}
	public String getFI() {
		return FI;
	}
	public void setFI(String fI) {
		FI = fI;
	}
	public String getDR() {
		return DR;
	}
	public void setDR(String dR) {
		DR = dR;
	}
	public String getKR() {
		return KR;
	}
	public void setKR(String kR) {
		KR = kR;
	}
	public String getMS() {
		return MS;
	}
	public void setMS(String mS) {
		MS = mS;
	}
	public String getFXFL() {
		return FXFL;
	}
	public void setFXFL(String fXFL) {
		FXFL = fXFL;
	}
	public String getBNCH() {
		return BNCH;
	}
	public void setBNCH(String bNCH) {
		BNCH = bNCH;
	}
	public String getGDNC() {
		return GDNC;
	}
	public void setGDNC(String gDNC) {
		GDNC = gDNC;
	}
	public String getSPRD() {
		return SPRD;
	}
	public void setSPRD(String sPRD) {
		SPRD = sPRD;
	}
	public String getCPN() {
		return CPN;
	}
	public void setCPN(String cPN) {
		CPN = cPN;
	}
	public String getYLD() {
		return YLD;
	}
	public void setYLD(String yLD) {
		YLD = yLD;
	}
	public String getPrice() {
		return Price;
	}
	public void setPrice(String price) {
		Price = price;
	}	
}
