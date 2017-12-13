package source;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.ApexPage;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.tooling.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadApexPageSync {
	
	private XSSFWorkbook workBook;
	
	public void readApexPage(String type,List<String> objectsList) throws Exception {
		Util.logger.info("ReadApexPage Start.");	
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String, String> resultMap = null;
		try {
			resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		} catch (ConnectionException e1) {
			e1.printStackTrace();
		}
		
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//Create catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		
		//Create ApexPage sheet
		String apSheetName = Util.makeSheetName("ApexPage");
		XSSFSheet excelApSheet= excelTemplate.createSheet(Util.cutSheetName(apSheetName));
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelApSheet,Util.cutSheetName(apSheetName),apSheetName);
		String names = ut.getObjectNames(objectsList);
		String sql2 = "Select  Name,NamespacePrefix,createdByID,createdDate,LastModifiedByID,LastModifiedDate From ApexPage WHERE Name in ("+ names +") Order By Name";
		SObject[] SObjects2= ut.apiQuery2(sql2);
		Map<String,com.sforce.soap.tooling.sobject.ApexPage> appMap = new HashMap<String,com.sforce.soap.tooling.sobject.ApexPage>();
		for(com.sforce.soap.tooling.sobject.SObject obj : SObjects2){
			com.sforce.soap.tooling.sobject.ApexPage app=(com.sforce.soap.tooling.sobject.ApexPage)obj;
			String keyStr="";
			if(app.getNamespacePrefix()!=null){
				keyStr=app.getNamespacePrefix()+"__";
			}
			appMap.put(keyStr+app.getName(), app);
		}
		List<String []> exportList = new ArrayList<String []>();
		//Create Apex Page Table		
		//ページ
		excelTemplate.createTableHeaders(excelApSheet,"Apex Page",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
		
		for (Metadata md : mdInfos) {
			if (md != null) {
				ApexPage obj = (ApexPage) md;    	
				int cellNum=1;
				XSSFRow columnRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				com.sforce.soap.tooling.sobject.ApexPage app=appMap.get(Util.nullFilter(obj.getFullName()));
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"ApexPage."+obj.getFullName()));
										
				}
				//名前
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getFullName()));
				//表示ラベル
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getLabel()));
				//名前空間プレフィックス
				if(app!=null){
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(app.getNamespacePrefix()));
				}else{
					excelTemplate.createCell(columnRow,cellNum++,"");
				}
				//APIバージョン
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getApiVersion()));
				//説明
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
				//モバイルアプリケーション使用可能
				excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(obj.getAvailableInTouch())));
				//GET要求のCSRF保護が必要
				excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(obj.getConfirmationTokenRequired())));
				if(app!=null){
					//作成者
					excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Id", Util.nullFilter(app.getCreatedById())));
					//作成日
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(app.getCreatedDate())));
					//最終更新者
					excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Id", Util.nullFilter(app.getLastModifiedById())));
					//最終更新日
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(app.getLastModifiedDate())));
				}else{
					excelTemplate.createCell(columnRow,cellNum++,"");
				}
				
				//Export source files
				String [] nameAndBody = new String[2];
				String s = new String(obj.getContent(), "UTF-8");
				nameAndBody[0] = String.valueOf(obj.getFullName()+".page");
				nameAndBody[1] = String.valueOf(s);
				exportList.add(nameAndBody);
			}
		}
		
		//Create Package Version Table
		/*
		  excelTemplate.createTableHeaders(excelApSheet,"Package Version",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
		
		
		for (Metadata md : mdInfos) {
			if (md != null) {
				ApexPage obj = (ApexPage) md;    	
				XSSFRow columnRowOne = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				if(obj.getPackageVersions().length>0){
					for(int i = 0;i<obj.getPackageVersions().length;i++){
						PackageVersion pkv = (PackageVersion)obj.getPackageVersions()[i];
						int cellNum=0;
						excelTemplate.createCell(columnRowOne,cellNum++,pkv.getNamespace());
						excelTemplate.createCell(columnRowOne,cellNum++,String.valueOf(pkv.getMajorNumber()));
						excelTemplate.createCell(columnRowOne,cellNum++,String.valueOf(pkv.getMinorNumber()));
					}
				}
			}
		}
		*/
		excelTemplate.adjustColumnWidth(excelApSheet);
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			excelTemplate.exportExcel(type,"");
			ut.exportSourceFile(type,exportList);
		}else{
			Util.logger.error("***no result to export!!!");
		}
		Util.logger.info("ReadApexPage End.");	
	}
}
