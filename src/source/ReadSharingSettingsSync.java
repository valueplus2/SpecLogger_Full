package source;


import java.io.IOException;
import java.net.URLDecoder;
import java.util.ArrayList;

import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wsc.MetadataLoginUtil;


import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.SharingCriteriaRule;
import com.sforce.soap.metadata.SharingOwnerRule;
import com.sforce.soap.metadata.SharingRules;
import com.sforce.ws.ConnectionException;

public class ReadSharingSettingsSync {

	/**
	 * @param args
	 */
	private XSSFWorkbook workBook;
	Util ut = new Util();
	public void readSharingSettings(String type,List<String> objectsList)throws Exception{
		Util.logger.info("ReadSharingRules Start.");
		List<Metadata> mdInfos = ut.readMateData("SharingRules",objectsList);
		//get all the objects
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<String> allFile = new ArrayList<String>();
		ListMetadataQuery query = new ListMetadataQuery();
		query.setType("CustomObject");
		FileProperties[] lmr = MetadataLoginUtil.metadataConnection.listMetadata(
				new ListMetadataQuery[] { query }, Util.API_VERSION);
		if (lmr != null) {
			for (FileProperties n : lmr) {
				allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
				Util.logger.debug("allFile="+allFile);
			}
		}		
		List<Metadata> mdIn = ut.readMateData("CustomObject",allFile);

		Map<String,String> sharingResultMap= ut.getComparedResult("SharingRules",UtilConnectionInfc.getLastUpdateTime());
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		//Organization-Wide Defaults
		//組織の共有設定
		if(objectsList.contains("Organization-Wide Defaults")){
			String sheetname=Util.makeSheetName("Organization_Wide_Defaults");
			XSSFSheet defaultSheet = excelTemplate.createSheet(Util.cutSheetName(sheetname));
			excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,defaultSheet,Util.cutSheetName(sheetname),sheetname);
			excelTemplate.createTableHeaders(defaultSheet, "Organization-Wide Defaults",defaultSheet.getLastRowNum()+Util.RowIntervalNum);		
			for (Metadata md : mdIn) {
				if (md != null) {
					CustomObject co= (CustomObject)md;
					if(co.getSharingModel()!=null){
						int cellNum=1;
						XSSFRow columnRow = defaultSheet.createRow(defaultSheet.getLastRowNum()+1);
						//オブジェクト
						excelTemplate.createCell(columnRow,cellNum++,ut.getLabelApi(Util.nullFilter(co.getFullName())));
						//デフォルトの内部アクセス権
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("SharingModel", co.getSharingModel().name())));
						//デフォルトの外部アクセス権
						String str="";
						if(co.getExternalSharingModel()!=null){
							str=String.valueOf(co.getExternalSharingModel().name());
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("SharingModel",str)));
					    }else{
					    	excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("SharingModel", co.getSharingModel().name())));
					    }
						//階層を使用したアクセス許可(機能ない)
						excelTemplate.createCell(columnRow,cellNum++,"");
					}
				}
				excelTemplate.adjustColumnWidth(defaultSheet);
			}	
		}
		for(Metadata md : mdInfos){
			if(md!=null){
				SharingRules sr = (SharingRules)md;
				String objectName =Util.makeSheetName(sr.getFullName());
				XSSFSheet excelSharingSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelSharingSettingsSheet,Util.cutSheetName(objectName),objectName);	
				//レコード所有者に基づくルール
				excelTemplate.createTableHeaders(excelSharingSettingsSheet, "OwnerSharingRule",excelSharingSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
				if(sr.getSharingOwnerRules().length>0){
					for( Integer t=0; t<sr.getSharingOwnerRules().length; t++ ){
						int cellNum=1;
						SharingOwnerRule aosr=(SharingOwnerRule)sr.getSharingOwnerRules()[t];
						XSSFRow columnRow = excelSharingSettingsSheet.createRow(excelSharingSettingsSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(sharingResultMap,"SharingOwnerRule."+sr.getFullName()+"."+aosr.getFullName()));
							
						}
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(aosr.getFullName()));
						//所有者の所属
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getSharedTo(aosr.getSharedFrom())));
						//共有先
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getSharedTo(aosr.getSharedTo())));
						//表示ラベル
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(aosr.getLabel()));					
						//説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(aosr.getDescription()));									
						//アクセス権
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("ShareAccessLevel",aosr.getAccessLevel())));					
					}
				}	
				//条件に基づくルール
				excelTemplate.createTableHeaders(excelSharingSettingsSheet, "CriteriaBasedSharingRule",excelSharingSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
				if(sr.getSharingCriteriaRules().length>0){
					for( Integer t=0; t<sr.getSharingCriteriaRules().length; t++ ){
						int cellNum=1;
						SharingCriteriaRule scr=(SharingCriteriaRule)sr.getSharingCriteriaRules()[t];
						XSSFRow columnRow = excelSharingSettingsSheet.createRow(excelSharingSettingsSheet.getLastRowNum()+1);
						
						if(UtilConnectionInfc.modifiedFlag){
							//変更あり
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(sharingResultMap.get("SharingCriteriaRule."+sr.getFullName()+"."+scr.getFullName()))));
						}					
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(scr.getFullName()));
						//共有先
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getSharedTo(scr.getSharedTo())));
						//表示ラベル
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(scr.getLabel()));					
						//説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(scr.getDescription()));									
						//アクセス権
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("ShareAccessLevel",scr.getAccessLevel())));										
						//条件
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getFilterItem(sr.getFullName(),scr.getCriteriaItems())));
						//条件ロジック
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(scr.getBooleanFilter()));
					}				
				}
				if(sr.getSharingTerritoryRules().length>0){
					//have not in use yet.
				}	
				excelTemplate.adjustColumnWidth(excelSharingSettingsSheet);
			}
		}
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}
		else{
			Util.logger.warn("***no result to export!!!");
		}
		Util.logger.info("ReadSharingRules End.");
	} 
}
				
				
				

