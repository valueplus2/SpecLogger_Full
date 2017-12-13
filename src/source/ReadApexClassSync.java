package source;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.tooling.sobject.ApexClass;
import com.sforce.soap.tooling.sobject.SObject;

public class ReadApexClassSync {
	
	private XSSFWorkbook workBook;
	
	public void readApexClass(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadApexClass Start.");
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//Create catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		Util ut = new Util();
		Map<String,String> resultMap = ut.getComparedResult(type, UtilConnectionInfc.lastUpdateTime);
		String names = ut.getObjectNames(objectsList);
		//Create query String
		//String sql = "Select  Name, NamespacePrefix,  LastModifiedDate, ApiVersion, Status, Body From ApexClass WHERE Name in ("+ names +")";
		//SObject [] SObjects= ut.apiQuery(sql);
		String sql2 = "Select  Name, NamespacePrefix,LengthWithoutComments,LastModifiedByID,LastModifiedDate, ApiVersion, Status, Body From ApexClass WHERE Name in ("+ names +") Order By Name";
		SObject[] SObjects2= ut.apiQuery2(sql2);
		Integer rowNum = Util.RowIntervalNum;
		//Create Sheets
		String sheetname=Util.makeSheetName("ApexClass");
		XSSFSheet excelSheet= excelTemplate.createSheet(Util.cutSheetName(sheetname));
		//Create Catalog menu
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelSheet,Util.cutSheetName(sheetname),sheetname);
		//Create TableHeaders
		//Apex クラス
		excelTemplate.createTableHeaders(excelSheet,"ApexClass",excelSheet.getLastRowNum()+rowNum);
		//rowNum += 2;
		List<String []> exportList = new ArrayList<String []>();
		for(SObject obj : SObjects2){
			ApexClass apc=(ApexClass)obj;
			//Create columnRow
			//XSSFRow columnRow = excelSheet.createRow(rowNum++);
			XSSFRow columnRow = excelSheet.createRow(excelSheet.getLastRowNum()+1);
			Integer cellNum = 1;
			//変更あり
			if(UtilConnectionInfc.modifiedFlag){
				excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"ApexClass."+apc.getName()));									
			}	
			//名前
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apc.getName()));
			//名前空間プレフィックス
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apc.getNamespacePrefix()));
			//サイズ（コメント除き）
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apc.getLengthWithoutComments()));
			//APIバージョン
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apc.getApiVersion()));
			//状況
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("APEXCODEUNITSTATUS",Util.nullFilter(apc.getStatus())));
			//最終更新者
			excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Id", Util.nullFilter(apc.getLastModifiedById())));
			//最終更新日
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(apc.getLastModifiedDate())));

			//Export source files
			String [] nameAndBody = new String[2];
			nameAndBody[0] = String.valueOf(apc.getName()+".cls");
			nameAndBody[1] = String.valueOf(apc.getBody());
			exportList.add(nameAndBody);
			
		}

		excelTemplate.adjustColumnWidth(excelSheet);
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			excelTemplate.exportExcel(type,"");
			ut.exportSourceFile(type, exportList);
		}else{
			Util.logger.error("***no result to export!!!");
		}
		Util.logger.info("ReadApexClass End.");
	}
	
}
