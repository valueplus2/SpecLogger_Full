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







import com.sforce.soap.metadata.Document;
import com.sforce.soap.metadata.DocumentFolder;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.ws.ConnectionException;

public class ReadDocumentSync {
	private Util ut;
	private MetadataConnection metadataConnection;
	private XSSFWorkbook workbook;
	private CreateExcelTemplate createExcelTemplate;
	//private XSSFSheet catalogSheet;
	private Map<String, String> resultMap;
	public void readDocumentFolder(String type, List<String> objectsList)
			throws Exception {
		Util.logger.info("ReadDocumentFolder Start."); 
		ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		metadataConnection = UtilConnectionInfc.getMetadataConnection();
		createExcelTemplate = new CreateExcelTemplate(type);
		workbook = createExcelTemplate.workBook;
		//catalogSheet = createExcelTemplate.createCatalogSheet();
		resultMap = this.getCompare("DocumentFolder",
				UtilConnectionInfc.getLastUpdateTime());
		XSSFSheet folderSheet = workbook.createSheet(Util.makeSheetName(Util.cutSheetName("DocumentFolder")));
		//Create Catalog menu
		createExcelTemplate.createCatalogMenu(createExcelTemplate.catalogSheet, folderSheet,Util.cutSheetName(folderSheet.getSheetName()),Util.makeSheetName("DocumentFolder"));
		List<Metadata> mdInfc = ut.readMateData("DocumentFolder", objectsList);
		//Create TableHeaders(DocumentFolder)
		createExcelTemplate.createTableHeaders(folderSheet, "DocumentFolder",
				folderSheet.getLastRowNum() + Util.RowIntervalNum);
		
		for (Metadata m : mdInfc) {
			if(m!=null){
				int cellNum=1;
				DocumentFolder df = (DocumentFolder) m;
				XSSFRow row = folderSheet
						.createRow(folderSheet.getLastRowNum() + 1);
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					createExcelTemplate.createCell(row,cellNum++,ut.getUpdateFlag(resultMap,"DocumentFolder." + m.getFullName()));
					
				}
				//一意の名前
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(df.getFullName()));
				//表示ラベル
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(df.getName()));
				//アクセスタイプ 
				createExcelTemplate.createCell(row,cellNum++,Util.getTranslate("ACCESSTYPE", Util.nullFilter(df.getAccessType().name()))); 
				//公開フォルダのアクセス権 
				createExcelTemplate.createCell(row,cellNum++,Util.getTranslate("FOLDERACCESS",Util.nullFilter(df.getPublicFolderAccess().name()))); 
				//共有先
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(ut.getSharedTo(df.getSharedTo())));
			}	
		}

		ListMetadataQuery listMetadataQuery = new ListMetadataQuery();
		for (String folder : objectsList) {
			listMetadataQuery.setType("Document");
			listMetadataQuery.setFolder(folder);
			try {
				FileProperties[] fileProperties = metadataConnection
						.listMetadata(
								new ListMetadataQuery[] { listMetadataQuery },
								31.0);
				List<String> list = new ArrayList<String>();
				for (FileProperties f : fileProperties) {
					
					list.add(f.getFullName());
				}
				resultMap = ut.getComparedResult(type,folder, UtilConnectionInfc.getLastUpdateTime());
				this.readDocument(folder,list);
			} catch (ConnectionException e) {
				e.printStackTrace();
			}
		}
		createExcelTemplate.adjustColumnWidth(folderSheet);
		if (workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null) {
			//createExcelTemplate.adjustColumnWidth(catalogSheet);
			createExcelTemplate.exportExcel(type,"");
		} else {
			System.out.println("***no result to export!!!");
		}
		Util.logger.info("ReadDocumentFolder End.");
	}

	//Create document Table
	public void readDocument(String folder,List<String> list) throws IOException {
		List<Metadata> mdInfc = ut.readMateData("Document", list);	
		Util.logger.info("ReadDocument Start."); 
		XSSFSheet dSheet = createExcelTemplate.createSheet(Util.makeSheetName(Util.cutSheetName(folder)));
		createExcelTemplate.createCatalogMenu(createExcelTemplate.catalogSheet, dSheet, dSheet.getSheetName());
		//ドキュメント
		createExcelTemplate.createTableHeaders(dSheet, "document", dSheet.getLastRowNum()+Util.RowIntervalNum);
		
		for (Metadata m : mdInfc) {
			Document d = (Document) m;
			if(m!=null){
				XSSFRow row = dSheet.createRow(dSheet.getLastRowNum() + 1);
				int cellNum=1;
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					createExcelTemplate.createCell(row,cellNum++,ut.getUpdateFlag(resultMap,"Document." + m.getFullName()));
					
				}
				//フルパス名
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getFullName()));
				//ドキュメント名
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getName()));
				//説明
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getDescription()));
				//キーワード
				createExcelTemplate.createCell(row,cellNum++,Util.nullFilter(d.getKeywords()));
				//社外秘フラグ
				createExcelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(d.getInternalUseOnly())));
				//外部参照可
				createExcelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(d.getPublic())));
				this.writeFile(d.getContent(), d.getFullName());
			}
		}
		createExcelTemplate.adjustColumnWidth(dSheet);
		Util.logger.info("ReadDocument End."); 
	}

	public void writeFile(byte[] fileContent, String fileName)
			throws IOException {
		String[] str = fileName.split("/");
		String filePath = UtilConnectionInfc.getDownloadPath() + "/Document/"
				+ str[0];
		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		OutputStream out = new FileOutputStream(filePath + "/" + str[1]);
		out.write(fileContent);
		out.flush();
		out.close();
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
