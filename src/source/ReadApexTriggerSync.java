package source;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;






//import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadApexTriggerSync {
  private XSSFWorkbook workBook;
  
  public void readApexTrigger(String type,List<String> objectsList)throws Exception{
	  Util.logger.info("ReadApexTrigger Start.");	
	  /*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//out = excelTemplate.out;
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();	 
		
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		Map<String,String> resultMap = ut.getComparedResult(type, UtilConnectionInfc.lastUpdateTime);
		String names = ut.getObjectNames(objectsList);
		System.out.println("names:"+names);
		
	    //String sql = "Select Name,ApiVersion,Status,Body,LastModifiedDate from ApexTrigger where Name in ("+names+")";
	    //SObject[] objs = ut.apiQuery(sql);
	    
	    String sql2 = "Select Name,NamespacePrefix,ApiVersion,Status,Body,TableEnumOrId,UsageBeforeUpdate,"
	    		+ "UsageBeforeInsert,UsageBeforeDelete,UsageAfterUpdate,UsageAfterUndelete,UsageAfterInsert,"
	    		+ "UsageAfterDelete,LastModifiedDate,LengthWithoutComments from ApexTrigger where Name in ("+names+") Order By Name";
	    com.sforce.soap.tooling.sobject.SObject[] objs2 = ut.apiQuery2(sql2);
	    
	    //create sheet
	    String sheetname = Util.makeSheetName(type);
	    XSSFSheet excelSheet = excelTemplate.createSheet(Util.cutSheetName(sheetname));
	    //create catelog
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelSheet,Util.cutSheetName(sheetname),sheetname);
	    //create tableHeader
	    Integer rowNum = Util.RowIntervalNum;
	    //Apex トリガ
	    excelTemplate.createTableHeaders(excelSheet, "Apex Trigger",excelSheet.getLastRowNum()+ rowNum);
	    //rowNum += 2;
	    //export list data
	    List<String []> exportList = new ArrayList<String []>();
	    //loop to write apex trigger attribute to excel
	    if(objs2.length>0){
	    	
	    	for(com.sforce.soap.tooling.sobject.SObject obj:objs2){
	    		com.sforce.soap.tooling.sobject.ApexTrigger apt=(com.sforce.soap.tooling.sobject.ApexTrigger)obj;
		    	//create columnRow
		    	//XSSFRow columnRow = excelSheet.createRow(rowNum++);
	    		XSSFRow columnRow = excelSheet.createRow(excelSheet.getLastRowNum()+1);
		    	int cellNum=1;
		    	//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"ApexTrigger."+apt.getName()));
										
				}
				//名前
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getName()));
		    	//名前空間プレフィックス
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getNamespacePrefix()));
		    	//オブジェクト名
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getTableEnumOrId()));
		    	//サイズ（コメント除き）
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getLengthWithoutComments()));
		    	//APIバージョン
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getApiVersion()));
		    	//状況
		    	excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("APEXCODEUNITSTATUS", Util.nullFilter(apt.getStatus())));
		    	//Before Insertトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageBeforeInsert()));
		    	//Before Deleteトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageBeforeDelete()));
		    	//Before Updateトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageBeforeUpdate()));
		    	//After Insertトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageAfterInsert()));
		    	//After Deleteトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageAfterDelete()));
		    	//After Updateトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageAfterUpdate()));
		    	//After Undeleteトリガ
		    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apt.getUsageAfterUndelete()));
		    	//最終更新日
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(apt.getLastModifiedDate())));		    	
		    	//export source files
		    	String[] nameAndBody = new String[2];
		    	nameAndBody[0] = String.valueOf(apt.getName()+".trigger");
		    	nameAndBody[1] = String.valueOf(apt.getBody());
		    	exportList.add(nameAndBody);
		    }
	    }
	    
	    excelTemplate.adjustColumnWidth(excelSheet);
	    if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			excelTemplate.exportExcel(type,"");
			ut.exportSourceFile(type, exportList);
		}else{
			Util.logger.error("***no result to export!!!");
		}
	  Util.logger.info("ReadApexTrigger End.");	
  }
}
