package source;


import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

import com.sforce.soap.metadata.ExternalDataSource;
import com.sforce.soap.metadata.Metadata;
import com.sforce.ws.ConnectionException;

public class ReadExternalDataSourceSync {
	private XSSFWorkbook workBook;
	
	public void readExternalDataSource(String type,List<String> objectsList) throws Exception {
		Util.logger.info("readExternalDataSources Start.");
		Util ut = new Util();
		try{
			List<Metadata> mdInfos = ut.readMateData(type, objectsList);
			Map<String, String> resultMap = null;
			try {
				resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
			} catch (ConnectionException e1) {
				e1.printStackTrace();
			}
			Util.nameSequence=0;
			Util.sheetSequence=0;
			/*** Get Excel template and create workBook(Common) ***/
			CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
			workBook = excelTemplate.workBook;
			
			//Create catalog sheet
			//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		
			for (Metadata md : mdInfos) {
				if (md != null) {
					ExternalDataSource eds = (ExternalDataSource)md;
					/*** Object Attribute ***/
					String objectName = Util.makeSheetName(eds.getFullName());
					XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
					//创建目录信息
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
					/*** Object Override ***/
					excelTemplate.createTableHeaders(excelObjectSheet,"Object Attribute",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					//Create columnRow
					XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
					Integer cellNum = 1;
					if(UtilConnectionInfc.modifiedFlag){
						excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"ExternalDataSource."+eds.getFullName()));
						
					}	
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(eds.getFullName()));
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(eds.getLabel()));	
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(eds.getType()));	
					
					/*** Parameters ***/
					excelTemplate.createTableHeaders(excelObjectSheet,"Parameters",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(eds.getCustomConfiguration()!=null){
						XSSFRow paraRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						Integer paraCell=1;
						excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(eds.getEndpoint()));
						try{
							JSONObject jsonObject = new JSONObject(eds.getCustomConfiguration());					
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("timeout")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("noIdMapping")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("pagination")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("inlineCountEnabled")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("requestCompression")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("searchEnabled")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("searchFunc")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("format")));
							excelTemplate.createCell(paraRow,paraCell++,Util.nullFilter(jsonObject.getString("compatibility")));
						}catch(JSONException e){
							e.getStackTrace();
						}
					}
					
					/*** Authentication ***/
					excelTemplate.createTableHeaders(excelObjectSheet,"Authentication",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					XSSFRow authRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
					Integer authCell=1;				
					excelTemplate.createCell(authRow,authCell++,Util.nullFilter(eds.getCertificate()));
					excelTemplate.createCell(authRow,authCell++,Util.getTranslate("ExternalPrincipalType",Util.nullFilter(eds.getPrincipalType())));
					excelTemplate.createCell(authRow,authCell++,Util.getTranslate("AuthenticationProtocol",Util.nullFilter(eds.getProtocol())));
					excelTemplate.createCell(authRow,authCell++,Util.nullFilter(eds.getUsername()));
					excelTemplate.createCell(authRow,authCell++,Util.nullFilter(eds.getPassword()));
					excelTemplate.createCell(authRow,authCell++,Util.nullFilter(eds.getAuthProvider()));
					excelTemplate.createCell(authRow,authCell++,Util.nullFilter(eds.getOauthScope()));
					
					excelTemplate.adjustColumnWidth(excelObjectSheet);	
				}else{
					Util.logger.warn("Empty metadata.");
				}
								
			}
			if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
				//excelTemplate.adjustColumnWidth(catalogSheet);
				excelTemplate.exportExcel(type,"");
			}else{
				Util.logger.warn("***no result to export!!!");
			}
			Util.logger.error("ReadCustomObject complete.");
		}catch(Exception e){			
			Util.logger.error("ReadCustomObject failure.");
			Util.logger.error("",e);
		}
		Util.logger.info("ReadCustomObject End.");	
	}

}
