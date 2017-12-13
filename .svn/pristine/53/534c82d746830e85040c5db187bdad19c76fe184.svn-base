package source;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import wsc.WSC;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;


public class CreateExcelTemplate {

	Util ut = new Util();
	PoiUtil pu = new PoiUtil();
	public XSSFWorkbook workBook;
	//XML Result Map<TableHeaderName,ColumnHeaderNameList>
	Map<String,List<String>> headerMap = new LinkedHashMap<String,List<String>>();
	List<String> headerList = new ArrayList<String>();
	public XSSFCellStyle cellStyle;	 
	public String type;
	public CreateExcelTemplate(String type) throws Exception{
		this.type = type;
		CreateWorkBook(type);		
		cellStyle=createCellStyle();
		createHyperLinkStyle();
		columnHeaderStyle=createCHeaderStyle();			

	}

	
	/**
	 * Create WorkBook
	 * @throws FileNotFoundException 
	 */
	//set font name
	public String fontName="ＭＳ Ｐゴシック";
	public XSSFFont cellfont;
	public XSSFSheet catalogSheet;
	public Integer wbNumber=1;
	//public Integer maxFileSize=400000;//byte;
	public Integer maxFileSize=20000;//byte;
	public void CreateWorkBook(String pathKey) throws Exception{	
		workBook = new XSSFWorkbook();
		cellfont=workBook.createFont();	
		cellStyle = workBook.createCellStyle();
		hyperLinkStyle = workBook.createCellStyle();
		hyperLinkStyle1 = workBook.createCellStyle();
		columnHeaderStyle = (XSSFCellStyle)workBook.createCellStyle();
		tableHeaderStyle = (XSSFCellStyle)workBook.createCellStyle();	
		catalogSheet = createCatalogSheet();	
		readXmlTemplate(pathKey);
		//style only create once		
		cellStyle=createCellStyle();		
		columnHeaderStyle=createCHeaderStyle();	                     
        createHyperLinkStyle();	
	}

	public void adjustColumnWidth(XSSFSheet sheet){
		int firstRow = sheet.getFirstRowNum();
		int lastRow = sheet.getLastRowNum();
		int maxColumns=1;
	    //To count max columns 
		for(int i = firstRow; i <= lastRow; i++){
			if(sheet.getRow(i)!=null){
				int curMaxColumns=sheet.getRow(i).getPhysicalNumberOfCells();
			    if(maxColumns<curMaxColumns){
			    	maxColumns=curMaxColumns;
			    }

		    }
		}
	    //First auto size all column's width
		for (int x = 1; x <= maxColumns; x++) {
	    	sheet.autoSizeColumn(x);
	    }	
	    //Then auto size all  cell's height
		for(int i = firstRow; i <= lastRow; i++){
			if(sheet.getRow(i)!=null){
		    	for (int x = 0; x <= sheet.getRow(i).getPhysicalNumberOfCells(); x++) {
			    	//sheet.autoSizeColumn(x);
			    	//sheet.setColumnWidth(x, 10000);
			    	if(sheet.getRow(i).getCell(x)!=null){
			    		pu.autoSizeRow(workBook, sheet, sheet.getRow(i), sheet.getRow(i).getCell(x), x);
			    	}
			    	
			    }
		    }
		}    	
	}

	public OutputStream CreateOutputStream(String pathKey) throws FileNotFoundException{
		//Create files downLoad path
		String keys=Util.getTranslate("MetadataType",pathKey);
		Util.logger.debug("readXmlTemplate keys="+keys);	
		String path = UtilConnectionInfc.getDownloadPath() + "\\" +keys+".xlsx";
		//Create WorkBook
		OutputStream out = new FileOutputStream(path);
		return out;
	}
	
	public void readXmlTemplate(String templateName){
		try{ 
			
			Util.logger.debug("readXmlTemplate start.");	
			String fileName = Util.getTemplateName(templateName) +"_"+ UtilConnectionInfc.getLanguage();
			File  xmlFile;
			if(WSC.isBatch){
				xmlFile = new File(".././common/xml/"+fileName+".xml"); 
			}else{
				xmlFile = new File("./common/xml/"+fileName+".xml");
			}
			Util.logger.debug("readXmlTemplate xmlFile="+xmlFile);	
			DocumentBuilderFactory  builderFactory =  DocumentBuilderFactory.newInstance();               
			DocumentBuilder   builder = builderFactory.newDocumentBuilder();               
			Document  doc = builder.parse(xmlFile);  
			if(doc.hasChildNodes()){
				//Create excel template
				printNode(doc.getChildNodes());
			}  
			Util.logger.debug("readXmlTemplate completed.");	
		}catch(Exception  e){  
			Util.logger.error(e.getStackTrace());	
			
		}	
	}
	
	
	public void printNode(NodeList nodeList){ 
		Util.logger.debug("printNode start.");
		String tableHeader = "";
		for(int i = 0;  i<nodeList.getLength(); i++){  
			Node  node = (Node)nodeList.item(i);
			if(node.getNodeType() == Node.ELEMENT_NODE){ 
				if( node.getNodeName().equals("tableHeader") ){
					tableHeader = node.getAttributes().getNamedItem("type").getNodeValue();
					headerList = new ArrayList<String>();
					headerList.add(node.getTextContent());
				} 
				if( node.getNodeName().equals("value") ){
					headerList.add(node.getTextContent());
				}
				if(node.hasChildNodes()){
					printNode(node.getChildNodes());  
				}  
			}
		}
		if( tableHeader != ""){
			headerMap.put(tableHeader,headerList);
		}
		Util.logger.debug("printNode completed.");
	}
	
	/**
	 * Create Catalog sheet
	 * 
	 * @return
	 * 		XSSFSheet
	 * @throws ConnectionException 
	 */
	public XSSFSheet createCatalogSheet(){
		Util.logger.debug("createCatalogSheet started.");
		System.out.println(Util.getTranslate("Common", "Index"));
		createCover();
		XSSFSheet catalogSheet = createSheet(Util.getTranslate("Common", "Index"));
		XSSFRow itemRow = catalogSheet.createRow(0);
		itemRow.createCell(0).setCellValue(Util.getTranslate("Common", "Index"));
		createCellName(Util.getTranslate("Common", "Index"), Util.getTranslate("Common", "Index"),1);
		itemRow = catalogSheet.createRow(1);
		itemRow.createCell(1).setCellValue(Util.getTranslate("Common", "COLUMNHEADRER"));
		itemRow.getCell(1).setCellStyle(columnHeaderStyle);
		Util.logger.debug("createCatalogSheet completed.");
		return catalogSheet;
	}
	
	private void createCover(){
		Util.logger.debug("createCover started.");
		String nowTimes = "";
		String orgName = "";
		Date dt = new Date();   
	    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");   
	    nowTimes=sdf.format(dt); 
		XSSFSheet coverSheet = createSheet(Util.getTranslate("Common", "cover"));
		
		try {
			String sql = "Select Name FROM Organization";
			SObject[] obj;
			obj = ut.apiQuery(sql);
			orgName = String.valueOf(obj[0].getField("Name"));
		} catch (ConnectionException e) {
			e.printStackTrace();
		}
		
		XSSFRow coverRow1 = coverSheet.createRow(1);
		XSSFCell cell1 = coverRow1.createCell(1);  
		cell1.setCellValue(Util.getTranslate("Common", "coverOrg")+" : " +orgName );
		cell1.setCellStyle(createTHeaderStyle());
		
		XSSFRow coverRow2 = coverSheet.createRow(2);
		XSSFCell cell2 = coverRow2.createCell(1);  
		cell2.setCellValue(Util.getTranslate("Common", "covertimeStamp")+" : " +nowTimes );
		cell2.setCellStyle(createTHeaderStyle());
		
		XSSFFont font = workBook.createFont();
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font.setFontHeightInPoints((short)30);
		XSSFCellStyle style = (XSSFCellStyle)workBook.createCellStyle();
		style.setFont(font);
		
		XSSFRow coverRow3 = coverSheet.createRow(10);
		XSSFCell cell3 = coverRow3.createCell(4);  
		cell3.setCellValue(Util.getTranslate("Common", "coverTitle"));
		cell3.setCellStyle(style);
		coverSheet.autoSizeColumn((short)1);
		coverSheet.autoSizeColumn((short)4);
		Util.logger.debug("createCover completed.");
	}
	
	
	/**
	 * 
	 * @param 
	 * 		XSSFSheet catalogSheet
	 * 		String excelSheet
	 * 		Integer sheetName
	 * @return 
	 * 		void
	 * 		
	 */
	public void createCatalogMenu(XSSFSheet catalogSheet,XSSFSheet excelSheet, String sheetName) {
		//Create catalog sheet values
		createlinkValueNoStyle(catalogSheet, catalogSheet.getLastRowNum() + 1, 1,sheetName,sheetName);
		//Create object sheet values
		createCellName(sheetName, sheetName, 1);
		createlinkValueNoStyle(excelSheet, 0, 0, Util.getTranslate("Common", "Index"), Util.getTranslate("Common", "BackToIndex"));
	}
	/**
	 * 
	 * @param 
	 * 		XSSFSheet catalogSheet
	 * 		String excelSheet
	 * 		String sheetName
	 * 		String displayName
	 * 
	 * @return 
	 * 		void
	 * 		
	 */
	public void createCatalogMenu(XSSFSheet catalogSheet,XSSFSheet excelSheet, String sheetName,String displayName) {
		//Create catalog sheet values
		createlinkValueNoStyle(catalogSheet, catalogSheet.getLastRowNum() + 1, 1,sheetName,displayName);
		//Create object sheet values
		createCellName(sheetName, sheetName, 1);
		createlinkValueNoStyle(excelSheet, 0, 0, Util.getTranslate("Common", "Index"), Util.getTranslate("Common", "BackToIndex"));
	}	
	/**
	 * 
	 * @param 
	 * 		XSSFSheet sheet
	 * 		String TableHeaderName
	 * 		Integer rowNum
	 * @return 
	 * 		void
	 * 		
	 */
	public void createTableHeaders(XSSFSheet sheet,String TableHeaderName,Integer rowNum){
		Util.logger.debug("createTableHeaders TableHeaderName="+TableHeaderName+"; rowNum="+rowNum);
		
		//add by cheng 15-11-19 start
		String str = Util.makeNameValue(TableHeaderName);
		Integer in1;
		if(sheet.getRow(2)!=null){
			 in1 = (int)sheet.getRow(2).getLastCellNum();
		}else{
			in1 = (int)sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
		}
		createlinkValueNoStyle(sheet, 2,in1, str, headerMap.get(TableHeaderName).get(0));
		createCellName(str,sheet.getSheetName(),rowNum+1);
		//add by cheng 15-11-19 end
		
		//Create tableHeaderRow
		XSSFRow tableHeaderRow = sheet.createRow(rowNum);
		//Create ColumnHeaderRow
		XSSFRow columnHeaderRow = sheet.createRow(rowNum+1);
		for(Integer i=0,cellNum = 1; i<headerMap.get(TableHeaderName).size(); i++){			 
			if(!UtilConnectionInfc.modifiedFlag&&(headerMap.get(TableHeaderName).get(i).contains("変更あり")||headerMap.get(TableHeaderName).get(i).contains("Changed"))){
				continue;
			}else{
				if(i==0){
					 XSSFCell cell = tableHeaderRow.createCell(0);  
					 cell.setCellValue(headerMap.get(TableHeaderName).get(i));
					 cell.setCellStyle(createTHeaderStyle());
				}else{
					XSSFCell cell = columnHeaderRow.createCell(cellNum);  
					cell.setCellValue(headerMap.get(TableHeaderName).get(i));
					cell.setCellStyle(columnHeaderStyle);
					cellNum++;
				}	
				
			}
		}
		
		/*
		for(short i=0;i<sheet.getRow(sheet.getLastRowNum()).getPhysicalNumberOfCells();i++){
			sheet.autoSizeColumn((short)i);
		}*/
	}
	
	/**
	 * 
	 * @param 
	 * 		String sheetName
	 * @return 
	 * 		XSSFSheet sheet
	 */
	public XSSFSheet createSheet(String sheetName){
		//add by cheng 2015-11-19 start
		XSSFSheet Xssf = workBook.createSheet(sheetName);
		Xssf.createRow(2);
		Xssf.getRow(2).createCell(0);
		//add by cheng 2015-11-19 end
		return Xssf;
		
	}
	
	/**
	 * 
	 * @param 
	 * 		String cellName
	 * 		String sheetName
	 * 		Integer rowNum
	 * @return 
	 * 		void
	 * 		
	 */
	public void createCellName(String cellName,String sheetName,Integer rowNum){
		//Name namedCell = workBook.createName();
		if(workBook.getName(cellName) == null){
			Name namedCell = workBook.createName();
			namedCell.setNameName(cellName);
			// cell reference
			String reference = sheetName + "!B" + rowNum; 
			namedCell.setRefersToFormula(reference);
		}
	}
	
	/**
	 * 
	 * @param 
	 * 		XSSFSheet excelSheet
	 * 		Integer rowNum
	 * 		Integer cellNum
	 * 		String	hyperVal
	 * 		String	displayVal
	 * @return 
	 * 		void
	 * 		
	 */
	public void createCellValue(XSSFSheet excelSheet,Integer rowNum,Integer cellNum,String hyperVal,String displayVal){
		if(excelSheet.getRow(rowNum)!=null){
			XSSFCell cell = excelSheet.getRow(rowNum).createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(hyperLinkStyle);
		}else{
			XSSFRow itemRow = excelSheet.createRow(rowNum);
			XSSFCell cell = itemRow.createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(hyperLinkStyle);
		}
		//excelSheet.autoSizeColumn((short)Short.parseShort(String.valueOf(cellNum)));
	}
	public void createlinkValue(XSSFSheet excelSheet,Integer rowNum,Integer cellNum,String hyperVal,String displayVal){
		if(excelSheet.getRow(rowNum)!=null){
			XSSFCell cell = excelSheet.getRow(rowNum).createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(createLinkStyle());
		}else{
			XSSFRow itemRow = excelSheet.createRow(rowNum);
			XSSFCell cell = itemRow.createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(createLinkStyle());
		}
	}
	//add by cheng 2016-2-3 start
	public void createlinkValueNoStyle(XSSFSheet excelSheet,Integer rowNum,Integer cellNum,String hyperVal,String displayVal){
		if(excelSheet.getRow(rowNum)!=null){
			XSSFCell cell = excelSheet.getRow(rowNum).createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(createHeadLinkStyle());
		}else{
			XSSFRow itemRow = excelSheet.createRow(rowNum);
			XSSFCell cell = itemRow.createCell(cellNum);  
			cell.setCellFormula("HYPERLINK(\"#"+hyperVal+"\",\""+displayVal+"\")");
			cell.setCellStyle(createHeadLinkStyle());
		}
	}
	//add by cheng 2016-2-3 end
	/**
	 * 
	 */
	public void createCell(XSSFRow row,Integer cellNum, String displayVal){
		if(displayVal==null){
			displayVal="";
		}
		XSSFCell cell = row.createCell(cellNum);
		cell.setCellValue(displayVal);
		cell.setCellStyle(cellStyle);	
	}

	public XSSFCellStyle createCellStyle(){
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setWrapText(true);
        /* Set the font name to MS PGothic */
        cellfont.setFontName(fontName);		
        cellStyle.setFont(cellfont);
		return cellStyle;
	}	
	/**
	 * 
	 * String Filter
	 * 
	 */
	public String stringFilter(String str){
		String regEx="[`~!@#$%^&*()+=|{}':;',\\-\\[\\]<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。，、？]";  
		Pattern pat = Pattern.compile(regEx);     
		Matcher mat = pat.matcher(str.replaceAll(" ", ""));     
		String ret = mat.replaceAll("").trim();  
		//sheetName length must less than 31
		if(ret.length()>31){
			ret =ret.substring(31) ;
		}
		return ret;
	}
	
	
	
	/**
	 * 
	 * Export method
	 * 
	 * @param 
	 * 		String type
	 * @return 
	 * 		void
	 */
	public void exportExcel(String fileName,String part) throws FileNotFoundException{
	   try {  		
  		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
  		workBook.write(byteArrayOutputStream);	
       } catch (Exception e) {
          Util.logger.error(e.getStackTrace());
       }		
		OutputStream out;			
		try {			
			//Create files downLoad path
			String keys=Util.getTranslate("MetadataType",fileName)+part;
			String path = UtilConnectionInfc.getDownloadPath() + "\\" +keys+".xlsx";
			//Create WorkBook
			out = new FileOutputStream(path);
			adjustColumnWidth(catalogSheet);
			workBook.write(out);
			out.flush();
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	
	/**
	 * 
	 * Create Excel Style
	 * 
	 */
	public XSSFCellStyle hyperLinkStyle;
	public XSSFCellStyle createHyperLinkStyle(){
		XSSFFont cellFont= workBook.createFont();
		cellFont.setUnderline((byte) 1);
		cellFont.setColor((short)10);
		cellFont.setFontName(fontName);
		hyperLinkStyle.setFont(cellFont);
		hyperLinkStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		hyperLinkStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		hyperLinkStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		hyperLinkStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		hyperLinkStyle.setWrapText(true);
		hyperLinkStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		return hyperLinkStyle;
	}
	public XSSFCellStyle createLinkStyle(){
		XSSFFont cellFont= workBook.createFont();
		cellFont.setUnderline((byte) 1);
		cellFont.setColor((short)10);
		cellFont.setFontName(fontName);	
		hyperLinkStyle.setWrapText(true);
		hyperLinkStyle.setFont(cellFont);			
		return hyperLinkStyle;
	}
	//add by cheng 2016-2-3 start
	public XSSFCellStyle hyperLinkStyle1;
	public XSSFCellStyle createHeadLinkStyle(){
		XSSFFont cellFont= workBook.createFont();
		cellFont.setUnderline((byte) 1);
		cellFont.setColor((short)10);
		cellFont.setFontName(fontName);	
		hyperLinkStyle1.setWrapText(true);
		hyperLinkStyle1.setFont(cellFont);			
		return hyperLinkStyle1;
	}
	//add by cheng 2016-2-3 end
	public XSSFCellStyle tableHeaderStyle;
	public XSSFCellStyle  createTHeaderStyle(){
		tableHeaderStyle.setFont(tHeaderFont());
		return tableHeaderStyle;
	}
	public XSSFCellStyle columnHeaderStyle;
	public XSSFCellStyle  createCHeaderStyle(){
		columnHeaderStyle.setFont(cHeaderFont());
		columnHeaderStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		columnHeaderStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		columnHeaderStyle.setFillForegroundColor((short)22);
		columnHeaderStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		columnHeaderStyle.setWrapText(true);
		columnHeaderStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		columnHeaderStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		columnHeaderStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		columnHeaderStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);		
		return columnHeaderStyle;
	}

	public XSSFCellStyle changeLineStyle(){
		XSSFCellStyle itemStyle = (XSSFCellStyle)workBook.createCellStyle();
		itemStyle.setWrapText(true);
		return itemStyle;
	}
	public Font tHeaderFont(){
		XSSFFont font = workBook.createFont();
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font.setFontHeightInPoints((short)10);
		font.setFontName(fontName);	
		return font;
	}
	
	public Font cHeaderFont(){
		XSSFFont font = workBook.createFont();
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font.setFontHeightInPoints((short)10);
		font.setColor((short)8);
		font.setFontName(fontName);	
		return font;
	}

}
