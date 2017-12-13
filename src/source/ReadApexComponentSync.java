package source;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.ApexComponent;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.PackageVersion;
import com.sforce.soap.tooling.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadApexComponentSync {
private XSSFWorkbook workBook;
	
	public void readApexComponent(String type,List<String> objectsList) throws Exception {
		Util.logger.info("readApexComponent Start.");
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String, String> resultMap = null;
		try {
			resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		} catch (ConnectionException e1) {
			Util.logger.error(e1.getMessage());
		}
		
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//create catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		
		//create ApexComponent's sheet
		String apSheetName =Util.makeSheetName("ApexComponent");
		XSSFSheet excelApSheet= excelTemplate.createSheet(Util.cutSheetName(apSheetName));
		//create catalog 
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelApSheet,Util.cutSheetName(apSheetName),apSheetName);
		
		List<String []> exportList = new ArrayList<String []>();
		//create Apex Component's Table
		String names = ut.getObjectNames(objectsList);
		//コンポーネント
		excelTemplate.createTableHeaders(excelApSheet,"Apex Component",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
		String sql2 = "Select  Name,NamespacePrefix,createdByID,createdDate,LastModifiedByID,LastModifiedDate From ApexComponent WHERE Name in ("+ names +") Order By Name";
		com.sforce.soap.tooling.sobject.SObject[] SObjects2= ut.apiQuery2(sql2);
		Map<String,com.sforce.soap.tooling.sobject.ApexComponent> apcMap = new HashMap<String,com.sforce.soap.tooling.sobject.ApexComponent>();
		for(com.sforce.soap.tooling.sobject.SObject obj : SObjects2){
			com.sforce.soap.tooling.sobject.ApexComponent apc=(com.sforce.soap.tooling.sobject.ApexComponent)obj;
			String keyStr="";
			if(apc.getNamespacePrefix()!=null){
				keyStr=apc.getNamespacePrefix()+"__";
			}
			apcMap.put(keyStr+apc.getName(), apc);
		}
		for (Metadata md : mdInfos) {
			if (md != null) {
				ApexComponent obj = (ApexComponent) md;    //create ApexComponent object	
				int cellNum=1;
				XSSFRow columnRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				com.sforce.soap.tooling.sobject.ApexComponent apc=apcMap.get(Util.nullFilter(obj.getFullName()));
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"ApexComponent."+obj.getFullName()));				
				}	
				//名前
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getFullName()));
				//表示ラベル
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getLabel()));
				//名前空間プレフィックス
				if(apc!=null){
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(apc.getNamespacePrefix()));
				}else{
					excelTemplate.createCell(columnRow,cellNum++,"");
				}
				//APIバージョン
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getApiVersion()));
				//説明
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
				if(apc!=null){
					//作成者
					excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Id", Util.nullFilter(apc.getCreatedById())));
					//作成日
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(apc.getCreatedDate())));
					//最終更新者
					excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Id", Util.nullFilter(apc.getLastModifiedById())));
					//最終更新日
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLocalTime(apc.getLastModifiedDate())));
				}else{
					excelTemplate.createCell(columnRow,cellNum++,"");
				}
			
				//Export source files
				String [] nameAndBody = new String[2];
				String s = new String(obj.getContent(), "UTF-8");
				nameAndBody[0] = String.valueOf(obj.getFullName()+".component");
				nameAndBody[1] = String.valueOf(s);
				exportList.add(nameAndBody);
			}
		}
		
		//create Package Version's Table
		//パッケージのバージョン
		excelTemplate.createTableHeaders(excelApSheet,"Package Version",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
		
		for (Metadata md : mdInfos) {
			if (md != null) {
				ApexComponent obj = (ApexComponent) md;    //create ApexComponent object	
				if(obj.getPackageVersions().length>0){
					int cellNum=1;
					for(int i = 0;i<obj.getPackageVersions().length;i++){
						PackageVersion pkv = (PackageVersion)obj.getPackageVersions()[i];
						XSSFRow columnRowOne = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
						//名前空間
						excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(pkv.getNamespace()));
						//メジャー番号
						excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(pkv.getMajorNumber()));
						//マイナー番号
						excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(pkv.getMinorNumber()));
						cellNum=1;
					}
				}
			}
		}
		excelTemplate.adjustColumnWidth(excelApSheet);
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			excelTemplate.exportExcel(type,"");
			ut.exportSourceFile(type,exportList);
		}else{
			Util.logger.error("***no result to export!!!");
		}
		Util.logger.info("readApexComponent End.");
	}
}

