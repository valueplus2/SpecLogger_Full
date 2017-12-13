package source;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;











import com.sforce.soap.metadata.AuraDefinitionBundle;
import com.sforce.soap.metadata.Document;
import com.sforce.soap.metadata.DocumentFolder;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.ws.ConnectionException;

public class ReadAuraDefinitionBundleSync {
	private Util ut;
	private MetadataConnection metadataConnection;
	private XSSFWorkbook workbook;
	private CreateExcelTemplate createExcelTemplate;
	//private XSSFSheet catalogSheet;
	private Map<String, String> resultMap;
	public void readAuraDefinitionBundle(String type, List<String> objectsList)
			throws Exception {
		Util.logger.info("ReadAuraDefinitionBundle Start."); 
		ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		metadataConnection = UtilConnectionInfc.getMetadataConnection();
		createExcelTemplate = new CreateExcelTemplate("AuraDefinitionBundle");
		workbook = createExcelTemplate.workBook;
		//catalogSheet = createExcelTemplate.createCatalogSheet();
		resultMap = this.getCompare("AuraDefinitionBundle",
				UtilConnectionInfc.getLastUpdateTime());
		List<Metadata> mdInfc = ut.readMateData("AuraDefinitionBundle", objectsList);	
		XSSFSheet dSheet = createExcelTemplate.createSheet(Util.makeSheetName(Util.cutSheetName("LightningComponent")));
		createExcelTemplate.createCatalogMenu(createExcelTemplate.catalogSheet, dSheet, dSheet.getSheetName());
		createExcelTemplate.createTableHeaders(dSheet, "AuraDefinitionBundle", dSheet.getLastRowNum()+Util.RowIntervalNum);
		for (Metadata m : mdInfc) {
			AuraDefinitionBundle d = (AuraDefinitionBundle) m;
			if(m!=null){
				XSSFRow row = dSheet.createRow(dSheet.getLastRowNum() + 1);
				int cellNum=1;
				if(UtilConnectionInfc.modifiedFlag){
					createExcelTemplate.createCell(row,cellNum++,ut.getUpdateFlag(resultMap,"AuraDefinitionBundle." + m.getFullName()));//変更あり
				}
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getFullName()));//名前
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getApiVersion()));//Apiバージョン
				createExcelTemplate.createCell(row,cellNum++,Util.getTranslate("AuraDefinitionBundleType",Util.nullFilter(d.getType())));//タイプ
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getDescription()));//説明
				this.writeFile(d.getControllerContent(), d.getFullName(),"Controller.js");
				this.writeFile(d.getDesignContent(), d.getFullName(),".design");
				this.writeFile(d.getDocumentationContent(), d.getFullName(),".auradoc");
				this.writeFile(d.getHelperContent(), d.getFullName(),"Helper.js");
				if(d.getType()!=null){
					this.writeFile(d.getMarkup(), d.getFullName(),Util.getTranslate("AuraDefinitionBundle",d.getType().name()));
				}
				this.writeFile(d.getRendererContent(), d.getFullName(),"Renderer.js");
				this.writeFile(d.getStyleContent(), d.getFullName(),".css");
				this.writeFile(d.getSVGContent(), d.getFullName(),".svg");
			}
		}
		createExcelTemplate.adjustColumnWidth(dSheet);
		if(workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//createExcelTemplate.adjustColumnWidth(catalogSheet);
			createExcelTemplate.exportExcel(type,"");
		}else{
			Util.logger.error("***no result to export!!!");
		}
		Util.logger.info("ReadAuraDefinitionBundle End."); 
	}

	public void writeFile(byte[] fileContent, String fileName,String type)
			throws IOException {
		//String[] str = fileName.split("/");
		if(fileContent!=null){
			String filePath = UtilConnectionInfc.getDownloadPath() + "/LightningComponent/"
					+ fileName;
			File file = new File(filePath);
			if (!file.exists()) {
				file.mkdirs();
			}
			OutputStream out = new FileOutputStream(filePath + "/" + fileName+type);
			out.write(fileContent);
			out.flush();
			out.close();
		}
	}

	public Map<String, String> getCompare(String type, Long lastUpdateTime)
			throws ConnectionException {
		ListMetadataQuery query = new ListMetadataQuery();
		query.setType(type);
		Map<String, String> map = new LinkedHashMap<String, String>();
		FileProperties[] filePro = metadataConnection.listMetadata(
				new ListMetadataQuery[] { query }, Util.API_VERSION);
		for (FileProperties fPro : filePro) {
			map.put(type + "." + fPro.getFullName(), fPro.getLastModifiedDate()
					.getTimeInMillis() > lastUpdateTime ? "TRUE" : "FALSE");
		}
		return map;
	}
}
