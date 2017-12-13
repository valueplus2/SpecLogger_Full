package source;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.PermissionSetApplicationVisibility;
import com.sforce.soap.metadata.PermissionSetApexClassAccess;
import com.sforce.soap.metadata.PermissionSetApexPageAccess;
import com.sforce.soap.metadata.PermissionSetCustomPermissions;
import com.sforce.soap.metadata.PermissionSetExternalDataSourceAccess;
import com.sforce.soap.metadata.PermissionSetFieldPermissions;
import com.sforce.soap.metadata.PermissionSetObjectPermissions;
import com.sforce.soap.metadata.PermissionSetRecordTypeVisibility;
import com.sforce.soap.metadata.PermissionSetUserPermission;
import com.sforce.soap.metadata.PermissionSetTabSetting;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.PermissionSet;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadPermissionSetSync {	
	private XSSFWorkbook workBook;
	public void readpermissionSet(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadPermissionSetSync Started.");	
		/*** Get Excel template and create workBook(Common) ***/
		Util ut = new Util();
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Util.sheetSequence=0;
		Util.nameSequence=0;
		//Create Catalog sheet
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());		
		//Map<String,String> layoutdMap = new HashMap<String,String>();
		for (Metadata md : mdInfos) {
			if (md != null) {
				PermissionSet perset = (PermissionSet) md;
				//Create Layout sheet
				String sheetName = Util.makeSheetName(perset.getLabel());
				XSSFSheet excelPermissionSetSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
				//目次を作成
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelPermissionSetSheet,Util.cutSheetName(sheetName),sheetName);
				int cellNum=1;
				//Create Layout Table
				//権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,type,excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				//Create layoutRow
				XSSFRow persetRow = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum()+1);
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(persetRow,cellNum++,ut.getUpdateFlag(resultMap,"PermissionSet." + perset.getFullName()));
					
				}
				//ラベル
				excelTemplate.createCell(persetRow,cellNum++,Util.nullFilter(perset.getLabel()));
				//説明
				excelTemplate.createCell(persetRow,cellNum++,Util.nullFilter(perset.getDescription()));
				//ユーザライセンス
				excelTemplate.createCell(persetRow,cellNum++,Util.nullFilter(perset.getLicense()));
				//String sql ="Select Id,Username From User Where UserType='"+String.valueOf(perset.getLabel())+"')"
				
				//カスタムアプリケーション権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetApplicationVisibility",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetApplicationVisibility psav :perset.getApplicationVisibilities()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetApplicationVisibility."+psav.getApplication()));
						//}
						//アプリケーション
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psav.getApplication()));
						//参照可能
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psav.getVisible()));
					}
				}
				
				//有効なApexクラス権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetApexClassAccess",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetApexClassAccess psac :perset.getClassAccesses()){
					if(perset.getClassAccesses()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetApexClassAccess."+psac.getApexClass()));
						//}
						//Apexクラス名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psac.getApexClass()));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psac.getEnabled()));
					}
				}
				//カスタム権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetCustomPermissions",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetCustomPermissions pscp :perset.getCustomPermissions()){
					if(perset.getCustomPermissions()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetCustomPermissions."+pscp.getName()));
						//}
						//権限名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(pscp.getName()));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(pscp.getEnabled()));
					}
				}	
				
				//有効なExternal Data Sourcesの権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetExternalDataSourceAccess",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetExternalDataSourceAccess psedsa :perset.getExternalDataSourceAccesses()){
					if(perset.getExternalDataSourceAccesses()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetExternalDataSourceAccess."+psedsa.getExternalDataSource()));
						//}
						//外部データソース
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psedsa.getExternalDataSource()));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psedsa.getEnabled()));
					}
				}	
				//項目の権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetFieldPermissions",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetFieldPermissions psfp :perset.getFieldPermissions()){
					if(perset.getFieldPermissions()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetFieldPermissions."+psfp.getField()));
						//}
						//項目
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psfp.getField()));
						//参照のみ
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psfp.getEditable()));
						//参照可能
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psfp.getReadable()));
					}
				}
				//オブジェクト権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetObjectPermissions",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetObjectPermissions psop :perset.getObjectPermissions()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetObjectPermissions."+psop.getObject()));
						//}
						//オブジェクト
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getObject()));
						//作成
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getAllowCreate()));
						//削除
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getAllowDelete()));
						//編集
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getAllowEdit()));
						//参照
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getAllowRead()));
						//すべて変更
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getModifyAllRecords()));
						//すべて表示
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psop.getViewAllRecords()));
					}
				}	
				//有効なApexクラスの権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetApexPageAccess",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetApexPageAccess psapa :perset.getPageAccesses()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetApexPageAccess."+psapa.getApexPage()));
						//}
						//Visualforce ページ名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psapa.getApexPage()));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psapa.getEnabled()));
					}
				}
				//レコードタイプの権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetRecordTypeVisibility",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetRecordTypeVisibility psrv :perset.getRecordTypeVisibilities()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetRecordTypeVisibility."+psrv.getRecordType()));
						//}
						//レコードタイプ
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psrv.getRecordType()));
						//選択済み
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psrv.getVisible()));
					}
				}
				//タブの権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetTabSetting",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetTabSetting psts :perset.getTabSettings()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetTabSetting."+psts.getTab()));
						//}
						//タブ
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psts.getTab()));
						//利用可能
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("PermissionSetTabVisibility.",Util.nullFilter(psts.getVisibility())));
					}
				}
				//一般ユーザ権限セット
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"PermissionSetUserPermission",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				for(PermissionSetUserPermission psup :perset.getUserPermissions()){
					if(perset.getApplicationVisibilities()!=null){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						//if(UtilConnectionInfc.modifiedFlag){
						//	excelTemplate.createCell(row,cellNum++,resultMap.get("PermissionSetUserPermission."+psup.getName()));
						//}
						//権限名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psup.getName()));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(psup.getEnabled()));
					}
					
				}
				//ユーザ名
				excelTemplate.createTableHeaders(excelPermissionSetSheet,"Username",excelPermissionSetSheet.getLastRowNum()+Util.RowIntervalNum);
				String sql ="Select Id,Name From UserLicense "+ "Where Name='"+perset.getLicense()+"'";
				//add by cheng 15-11-20 start
				sql += " order by id ";
				//add by cheng 15-11-20 end
				SObject [] SObjects= ut.apiQuery(sql);
				for(SObject obj:SObjects){
					String sql2 ="Select FirstName,LastName From User Where ProfileId in(Select Id From Profile Where UserLicenseId='"+obj.getField("Id")+"') And IsActive=True";
					//add by cheng 15-11-20 start
					sql2 += " order by id ";
					//add by cheng 15-11-20 end
					SObject [] Object= ut.apiQuery(sql2);
					for(SObject ob:Object){
						cellNum=1;
						XSSFRow row = excelPermissionSetSheet.createRow(excelPermissionSetSheet.getLastRowNum() + 1);
						String fullname=String.valueOf(ob.getField("LastName"))+" "+String.valueOf(ob.getField("FirstName"));
						//名前
						excelTemplate.createCell(row,cellNum,Util.nullFilter(fullname));
					}
				}
				excelTemplate.adjustColumnWidth(excelPermissionSetSheet);
			}
			
		}
		Util.logger.info("ReadPermissionSetSync End.");
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			System.out.println("***no result to export!!!");
		}		
	}
}