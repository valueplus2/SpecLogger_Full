package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.CustomMetadata;
import com.sforce.soap.metadata.CustomMetadataValue;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.CustomField;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.ReadResult;
import com.sforce.soap.metadata.ValidationRule;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

import wsc.MetadataLoginUtil;

public class ReadCustomMetadataSync {

	private XSSFWorkbook workBook;
	private Map<String,String> apiToLabelMap = new HashMap<String,String>();
	
	public void ReadCustomMetadata(String type,List<String> objectsList) throws IOException, ConnectionException{
		try{
			Util.logger.info("ReadCustomMetadata Start.");	
			Util ut = new Util();
			List<Metadata> mdInfos = ut.readMateData("CustomMetadata", objectsList);
			
			Map<String,List<CustomMetadata>> recordMap = new HashMap<String,List<CustomMetadata>>();
			//custommetadata structure is different with other components, so it needs to be queried one more time 
			ListMetadataQuery query = new ListMetadataQuery();
			query.setType("CustomMetadata");
			FileProperties[] lmr = MetadataLoginUtil.metadataConnection.listMetadata(
					new ListMetadataQuery[] { query }, Util.API_VERSION);
			if (lmr != null) {
				List<String> allFile = new ArrayList<String>();
				for (FileProperties n : lmr) {
					allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));					
				}
				String[] read = new String[allFile.size()];
				allFile.toArray(read);				
				ReadResult readResult = MetadataLoginUtil.metadataConnection.readMetadata(type, read);
				Metadata[] records = readResult.getRecords();
				for(Metadata md : mdInfos){
					CustomObject obj = (CustomObject) md;
					List<CustomMetadata> cmList = new ArrayList<CustomMetadata>();
					String typeName = obj.getFullName().replace("__mdt","");					
					for(Metadata r : records){
						CustomMetadata cm = (CustomMetadata)r;
						if(cm.getFullName().contains(typeName)){
							cmList.add(cm);
						}
					}
					recordMap.put(typeName, cmList);
				}
			}				
			
			Map<String,String> resultMap = ut.getComparedResult("CustomObject",UtilConnectionInfc.getLastUpdateTime());
			Util.nameSequence=0;
			Util.sheetSequence=0;
			/*** Get Excel template and create workBook(Common) ***/
			CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
			workBook = excelTemplate.workBook;
			//Create catalog sheet
			//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
			/*** Loop MetaData results ***/
			for (Metadata md : mdInfos) {
				if (md != null) {
					// Create CustomSetting object
					CustomObject obj = (CustomObject) md;

					if(obj.getVisibility() !=null){
						/**---create Custom Metadata Attribute table --*/
						String objectName = Util.makeSheetName(obj.getFullName());
						XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
						
						excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
						int cellNum = 1;
						//カスタムメタデータ型の詳細
						excelTemplate.createTableHeaders(excelObjectSheet,"Custom Metadata Attribute",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						apiToLabelMap = new HashMap<String,String>();
						if(obj.getFields().length>0){					
							for(CustomField cf:obj.getFields()){
								apiToLabelMap.put(cf.getFullName(), cf.getLabel());
							}
						}						
						
						//Create columnRow
						XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						//表示ラベル
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getLabel()));
						//オブジェクト名
						String typeName = obj.getFullName().replace("__mdt","");	
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(typeName));						
						//API 参照名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getFullName()));
						//説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
						//表示
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getVisibility()));
						
						/** -----create CustomField table --*/
						Map<String,String> fieldLabelMap = new HashMap<String,String>();
						//カスタム項目
						excelTemplate.createTableHeaders(excelObjectSheet,"Custom Field",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						CustomField cfList[] = obj.getFields();
						if(cfList!=null){
							for(CustomField cf:cfList){
								cellNum = 1;
								columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//表示ラベル
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getLabel()));
								//項目名
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getFullName()));
								//get field label to map
								fieldLabelMap.put(cf.getFullName(),cf.getLabel());
								//データ型
								excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("FIELDTYPE",Util.nullFilter(cf.getType())));
								//大文字と小文字を区別する
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE", Util.nullFilter(cf.getCaseSensitive()))));
								//デフォルト値
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDefaultValue()));
								//説明
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDescription()));
								//インラインヘルプテキスト
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getInlineHelpText()));
								//暗号化
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getEncrypted()))));
								//数式
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getFormula()));
								//数式内の空白の処理
								excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("TreatBlanksAs", Util.nullFilter(cf.getFormulaTreatBlanksAs())));
								//表示形式
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDisplayFormat()));
								//外部 ID
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getExternalId()))));
								//文字数
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getLength()));
								//小数点の位置
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getScale()));
								//精度
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getPrecision()));
								//必須項目
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getRequired()))));
								//ユニーク
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getUnique()))));
								//表示される行
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getVisibleLines()));
								//開始番号
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getStartingNumber()));
							}
						}						

						/*** Validation Rule***/
						//创建Validation Rule的Table(入力規則)
						excelTemplate.createTableHeaders(excelObjectSheet,"Validation Rule",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						if(obj.getValidationRules().length>0){
							for(int i=0;i<obj.getValidationRules().length;i++){
								XSSFRow validRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								ValidationRule vr = obj.getValidationRules()[i];
								cellNum=1;
								Util.logger.debug("ValidationRule="+vr);
								//変更あり
								if(UtilConnectionInfc.modifiedFlag){
									excelTemplate.createCell(validRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("ValidationRule."+obj.getFullName()+"."+vr.getFullName()))));
								}
								//ルール名
								excelTemplate.createCell(validRow,cellNum++,Util.nullFilter(vr.getFullName()));
								//有効
								excelTemplate.createCell(validRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(vr.getActive())));
								//説明
								excelTemplate.createCell(validRow,cellNum++,Util.nullFilter(vr.getDescription()));
								//エラー条件数式
								excelTemplate.createCell(validRow,cellNum++,Util.nullFilter(vr.getErrorConditionFormula()));
								//エラー表示場所
								if(vr.getErrorDisplayField()==null){
									excelTemplate.createCell(validRow,cellNum++,Util.getTranslate("ERRORMESSAGE", "DISPLAYDEFAULT"));
								}else{
									excelTemplate.createCell(validRow,cellNum++,this.apiToLabelapi(Util.nullFilter(vr.getErrorDisplayField())));
								}
								//エラーメッセージ	
								excelTemplate.createCell(validRow,cellNum++,Util.nullFilter(vr.getErrorMessage()));						
		
							}
						}
								
						
						/** -----create Manage table -----*/
						//管理
						excelTemplate.createTableHeaders(excelObjectSheet,"Manage",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow headerRow = excelObjectSheet.getRow(excelObjectSheet.getLastRowNum());
						
						if(recordMap.get(typeName)!=null){
							List<CustomMetadata> cmList = recordMap.get(typeName);														
							// write manage table header to excel
							XSSFCell cell = headerRow.createCell(headerRow.getLastCellNum());
							cell.setCellValue(typeName+"名");
							cell.setCellStyle(excelTemplate.createCHeaderStyle());								
							boolean firstHeader = true;						
							for(int i=0;i<cmList.size();i++){							
								//write custom setting records to excel
								cellNum = 1;
								columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//表示ラベル
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cmList.get(i).getLabel()));
								//保護コンポーネント
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cmList.get(i).getProtected()));
								//名
								String name = cmList.get(i).getFullName().replace(".md", "");
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(name.substring(name.lastIndexOf('.')+1)));
								
								//フィールド項目								
								for(CustomMetadataValue cmv : cmList.get(i).getValues()){									
									//header
									if(firstHeader){										
										XSSFCell cellfield = headerRow.createCell(headerRow.getLastCellNum());
										cellfield.setCellValue(fieldLabelMap.get(cmv.getField()));
										cellfield.setCellStyle(excelTemplate.createCHeaderStyle());
									}
									//record
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cmv.getValue()));
								}
								firstHeader = false;//header only need to repeat once
							}	

						}
																	
						/**----------------------------------------------------------*/
						excelTemplate.adjustColumnWidth(excelObjectSheet);
					} 

				}
			}
			if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
				//excelTemplate.adjustColumnWidth(catalogSheet);
				excelTemplate.exportExcel(type,"");
			}else{
				Util.logger.warn("***no result to export!!!");
			}
		}catch(Exception e){

		}
		Util.logger.info("ReadCustomMetadata End.");
	}
	public String apiToLabelapi(String apiName){
		String result="";
		if(apiToLabelMap.get(apiName)!=null){
			result=apiToLabelMap.get(apiName)+"("+apiName+")";
		}else{
			result=apiName+"("+apiName+")";
		}
		return result;
	}	
}

