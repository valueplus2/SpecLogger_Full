package source;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.ObjectRelationship;
import com.sforce.soap.metadata.ReportLayoutSection;
import com.sforce.soap.metadata.ReportType;
import com.sforce.soap.metadata.ReportTypeColumn;
import com.sforce.ws.ConnectionException;

public class ReadReportTypeSync {
   private XSSFWorkbook workbook;
   public void readReportType(String type,List<String> objectsList)throws Exception{
	   Util.logger.info("ReadReportType Start.");
	   Util ut = new Util();
	   List<Metadata> mdInfos = ut.readMateData(type, objectsList);
	   Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.lastUpdateTime);
	   Util.nameSequence=0;
		Util.sheetSequence=0;
	   /*** Get Excel template and create workBook(Common) ***/
	   CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
	   workbook = excelTemplate.workBook;
	   //out = excelTemplate.out;
       //XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();	   
	   /*** Loop MetaData results ***/
	   for(Metadata md:mdInfos){
		   if(md!=null){
			   //create a ReportType object
			   ReportType rt = (ReportType)md;
			   
			  /********Report Types attribute*********/
			  //create Report Type's sheet
			   //レポートタイプ
			  String sheetName = Util.makeSheetName(rt.getFullName());
			  XSSFSheet reporttypeSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
			//create catalog message
			 excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,reporttypeSheet,Util.cutSheetName(sheetName),sheetName);
			  //create report type table
			  excelTemplate.createTableHeaders(reporttypeSheet,"Report Type",reporttypeSheet.getLastRowNum()+Util.RowIntervalNum);

			  Integer rowNum = reporttypeSheet.getLastRowNum()+1;
			  //create a new row to write 
			  XSSFRow newRow = reporttypeSheet.createRow(rowNum);
			  //create cell and write in data
			  int cellNum = 1;
			  if(UtilConnectionInfc.modifiedFlag){
				    excelTemplate.createCell(newRow,cellNum++,ut.getUpdateFlag(resultMap,"ReportType." + rt.getFullName()));
					
			  }
			 //レポートタイプ名
			  excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rt.getFullName()));
			 //主オブジェクト
			  excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rt.getBaseObject()));
			//レポートタイプの表示ラベル
			  excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rt.getLabel()));
			 //説明
			  excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rt.getDescription()));
			//カテゴリに格納
			  excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("ReportTypeCategory", Util.nullFilter(rt.getCategory())));
			//リリース状況
			  excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(rt.getDeployed())));
			 //自動生成
			  excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(rt.isAutogenerated())));
			 
			  //create ObjectRelationship table 
			  //オブジェクトリレーションシップ
			  Integer orRow=reporttypeSheet.getLastRowNum()+Util.RowIntervalNum;
			  excelTemplate.createTableHeaders(reporttypeSheet, "Object Relationship", orRow);
			  ObjectRelationship objRelation=rt.getJoin();
			  while(objRelation!=null){
				  //create row
				  cellNum = 1;
				  XSSFRow relationRow = reporttypeSheet.createRow(reporttypeSheet.getLastRowNum()+1);
				//リレーションシップ
				  excelTemplate.createCell(relationRow,cellNum++,Util.nullFilter(objRelation.getRelationship()));
				 //外部結合
				  excelTemplate.createCell(relationRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(objRelation.getOuterJoin())));
				  objRelation = objRelation.getJoin(); //Recursive output 
			  }
			  //create a report layout section table
			  //レポートレイアウトセクション
			  Integer lastRow = reporttypeSheet.getLastRowNum()+Util.RowIntervalNum;
			  excelTemplate.createTableHeaders(reporttypeSheet, "Report Layout Section", lastRow);
			  ReportLayoutSection[] rs = rt.getSections();
			  //loop the sections 
			  for(Integer i = 0;i<rs.length;i++){
				  //create a new row
				  cellNum = 1;
				  XSSFRow sectionRow = reporttypeSheet.createRow(reporttypeSheet.getLastRowNum()+1);
				  ReportLayoutSection rls = rs[i];
				  excelTemplate.createCell(sectionRow,cellNum++,Util.nullFilter(rls.getMasterLabel()));
				  ReportTypeColumn[] rtc=rls.getColumns();
				  for(Integer j=0;j<rtc.length;j++){
					 if(j==0){
						//参照先
						 excelTemplate.createCell(sectionRow,cellNum++,Util.nullFilter(rtc[j].getTable()));
						//アイテム名
						 excelTemplate.createCell(sectionRow,cellNum++,Util.nullFilter(rtc[j].getField()));
						//表示名
						 excelTemplate.createCell(sectionRow,cellNum++,Util.nullFilter(rtc[j].getDisplayNameOverride()));
						//デフォルトで選択
						 excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(rtc[j].getCheckedByDefault())));
					 }else{
						 //create a new row
						 XSSFRow newSectionRow = reporttypeSheet.createRow(reporttypeSheet.getLastRowNum()+1);
						//
						 excelTemplate.createCell(newSectionRow,cellNum++,"");
						//参照先
						 excelTemplate.createCell(newSectionRow,cellNum++,Util.nullFilter(rtc[j].getTable()));
						//アイテム名
						 excelTemplate.createCell(newSectionRow,cellNum++,Util.nullFilter(rtc[j].getField()));
						//表示名
						 excelTemplate.createCell(newSectionRow,cellNum++,Util.nullFilter(rtc[j].getDisplayNameOverride()));
						//デフォルトで選択
						 excelTemplate.createCell(newSectionRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(rtc[j].getCheckedByDefault())));
					 }
					 cellNum = 1;
				  }
			  }
			  excelTemplate.adjustColumnWidth(reporttypeSheet);
		   }else{
			   Util.logger.warn("Empty metadata.");
		   }
	   }
	   if(workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
		   //excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("***no result to export!!!");
	  }
	   Util.logger.info("ReadReportType End."); 
   }
}