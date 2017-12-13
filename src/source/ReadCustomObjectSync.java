package source;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.ActionOverride;
import com.sforce.soap.metadata.BusinessProcess;
import com.sforce.soap.metadata.CompactLayout;
import com.sforce.soap.metadata.CustomField;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.CustomValue;
import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.ListView;
import com.sforce.soap.metadata.ListViewFilter;
import com.sforce.soap.metadata.LookupFilter;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.PicklistValue;
import com.sforce.soap.metadata.RecordType;
import com.sforce.soap.metadata.RecordTypePicklistValue;
import com.sforce.soap.metadata.SearchLayouts;
import com.sforce.soap.metadata.SharingReason;
import com.sforce.soap.metadata.SharingRecalculation;
import com.sforce.soap.metadata.ValidationRule;
import com.sforce.soap.metadata.ValueSet;
import com.sforce.soap.metadata.ValueSettings;
import com.sforce.soap.metadata.WebLink;
import com.sforce.ws.ConnectionException;

public class ReadCustomObjectSync {

	private XSSFWorkbook workBook;
	private Util util = new Util();
	private Map<String,String> apiToLabelMap = new HashMap<String,String>();
	//to avoid hyper link name exceed 32 character

	public void readCustomObject(String type,List<String> objectsList) throws IOException, ConnectionException{
		Util.logger.info("ReadCustomObject End.");
		try{
			Util.logger.info("ReadCustomObject Start.");
			//For each Excel file reset the count of hyper links to zero
			Util.nameSequence=0;
			Util.sheetSequence=0;
			
			Util ut = new Util();
			List<Metadata> mdInfos = ut.readMateData(type, objectsList);
			Map<String,String> resultMap = ut.getComparedResult("CustomObject",UtilConnectionInfc.getLastUpdateTime());
			/*** Get Excel template and create workBook(Common) ***/
			CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
			workBook = excelTemplate.workBook;
			//创建目录sheet
			//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
	
			/*** Loop MetaData results ***/
			Integer lastIndex=0;
			for (Metadata md : mdInfos) {
				lastIndex+=1;
				if (md != null) {
					// Create CustomObject object
					CustomObject obj = (CustomObject) md;
					Util.logger.debug("CustomObject="+obj);	
					/*** Object Attribute ***/
					String objectName = Util.makeSheetName(obj.getFullName());
					XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
					//创建目录信息
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
					
					//add by cheng 2015-11-19 start
					//excelObjectSheet.createRow(2);
					//excelObjectSheet.getRow(2).createCell(0);
					//add by cheng 2015-11-19 end
					
					//オブジェクト定義の詳細
					excelTemplate.createTableHeaders(excelObjectSheet,"Object Attribute",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					apiToLabelMap = new HashMap<String,String>();
					if(obj.getFields().length>0){					
						for(CustomField cf:obj.getFields()){
							apiToLabelMap.put(cf.getFullName(), cf.getLabel());
						}
					}
					//Create columnRow
					XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
					Integer cellNum = 1;
					//変更あり
					if(UtilConnectionInfc.modifiedFlag){	
						excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"CustomObject."+obj.getFullName()));
					}
					//API参照名
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getFullName()));
					//表示ラベル
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getLabel()));
					//表示ラベル(複数形)
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getPluralLabel()));
					//ディビジョンを有効化
					excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE", Util.nullFilter(obj.getEnableDivisions()))));
					//コンパクトレイアウト
					String defaultCompactLayout=Util.nullFilter(obj.getCompactLayoutAssignment());
					if(defaultCompactLayout.equals("SYSTEM")){
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("COMPACTLAYOUT","SYSTEM"));
					}else if(obj.getCompactLayouts().length>0){
						Boolean hasResult=false;
						for(CompactLayout ct:obj.getCompactLayouts()){
							if(ct.getFullName().equals(defaultCompactLayout)){
								excelTemplate.createCell(columnRow,cellNum++,ct.getLabel());
								hasResult=true;
								break;
							}
						}
						if(!hasResult){
							excelTemplate.createCell(columnRow,cellNum++,"");
						}
					}else{
						excelTemplate.createCell(columnRow,cellNum++,"");
					}
					//excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getCompactLayoutAssignment()));//コンパクトレイアウト
					String customHelp="";
					if(obj.getCustomHelp()!=null){
						customHelp=String.valueOf(obj.getCustomHelp())+"("+Util.getTranslate("Common","Scontrol")+")";
					}
					if(obj.getCustomHelpPage()!=null){
						customHelp=String.valueOf(obj.getCustomHelpPage())+"("+Util.getTranslate("Common","Visualforce")+")";
					}	
					//カスタムヘルプの設定
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(customHelp));
					if(type=="CustomObject"){
						//リリース状況
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("DeploymentStatus",Util.nullFilter(obj.getDeploymentStatus())));
					}else{
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE","N/A"));

					};
					//説明
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
					//活動を許可
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableActivities())));
					//高度なルックアップ
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableEnhancedLookup())));
					//フィード追跡
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableFeeds())));
					//項目履歴管理
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableHistory())));
					//レポートを許可
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableReports())));
					//共有を許可
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableSharing())));
					//BulkAPIアクセスを許可
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableBulkApi())));
					//ストリーミングAPIアクセス
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEnableStreamingApi())));	
					//String gender="";
					//if(obj.getGender() != null){
					//	gender=String.valueOf(obj.getGender());
					//}
					//Util.logger.debug("------gender-----"+);
					//gender
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getGender()));
					//household		
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getHousehold())));		
					if(obj.getNameField()!=null){
						
						if(obj.getNameField().getType().toString().equals("Text")){
							//レコード名
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getNameField().getLabel()));
						}else{
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getNameField().getDisplayFormat()));
						}
					}else{
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter("Name"));
					}
					//レコードタイプフィード追跡
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getRecordTypeTrackFeedHistory())));
					//レコードタイプ履歴追跡	
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getRecordTypeTrackHistory())));		
					//開始文字
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getStartsWith()));
					//Chatter グループ内で許可
					excelTemplate.createCell(columnRow, cellNum++, Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getAllowInChatterGroups())));
					
					/*** Action Override ***/
					//オーバーライドアクション
					excelTemplate.createTableHeaders(excelObjectSheet,"Action Override",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);	
					if(obj.getActionOverrides().length>0){
						for( Integer i=0; i<obj.getActionOverrides().length; i++){
							ActionOverride ac =(ActionOverride)obj.getActionOverrides()[i];
							Util.logger.debug("ActionOverride="+ac);
							if(!(String.valueOf(ac.getType()).equals("Default")||String.valueOf(ac.getType()).equals("Standard"))){
								XSSFRow actionRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								cellNum = 1;
								//表示ラベル
								excelTemplate.createCell(actionRow,cellNum++,Util.getTranslate("ActionName",Util.nullFilter(ac.getActionName())));
								//名前
								excelTemplate.createCell(actionRow,cellNum++,Util.nullFilter(ac.getActionName()));
								//説明
								excelTemplate.createCell(actionRow,cellNum++,Util.nullFilter(ac.getComment()));
								//内容のソース
								excelTemplate.createCell(actionRow,cellNum++,ac.getContent()+"("+Util.getTranslate("ActionOverrideType",Util.nullFilter(ac.getType()))+")");
								String tmpSkipRec=Util.getTranslate("BooleanValue","NULL");
								if(ac.getActionName()=="New"){
									tmpSkipRec=Util.getTranslate("BooleanValue",String.valueOf(ac.getSkipRecordTypeSelect()));
								}
								//レコードタイプ選択をスキップする
								excelTemplate.createCell(actionRow,cellNum++,Util.nullFilter(tmpSkipRec));
							}
						}
					}
	
	
					/*** Business Process ***/
					List<String> objWithBusinessProcess = Arrays.asList("Opportunity","Case","Solution","Lead");
					if(objWithBusinessProcess.contains(obj.getFullName())){
						//ビジネスプロセス
						excelTemplate.createTableHeaders(excelObjectSheet,"Business Process",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);	
						if(obj.getBusinessProcesses().length>0){
							for( Integer i=0; i<obj.getBusinessProcesses().length; i++){
								XSSFRow businessRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								BusinessProcess bp =(BusinessProcess)obj.getBusinessProcesses()[i];
								cellNum = 1;
								Util.logger.debug("BusinessProcess="+bp);
								//変更あり
								if(UtilConnectionInfc.modifiedFlag){
									excelTemplate.createCell(businessRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("BusinessProcess."+obj.getFullName()+"."+bp.getFullName()))));
								}
								//API名
								excelTemplate.createCell(businessRow,cellNum++,Util.nullFilter(bp.getFullName()));
								//有効
								excelTemplate.createCell(businessRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(bp.getIsActive())));
								//説明
								excelTemplate.createCell(businessRow,cellNum++,Util.nullFilter(bp.getDescription()));
								excelTemplate.createCell(businessRow,cellNum++,Util.getTranslate("Common",Util.nullFilter("None")));
								String value="";				
								for(PicklistValue pv:bp.getValues()){
									if(pv.getDefault()){
										value+=pv.getFullName()+"("+Util.getTranslate("Common","Default")+")\n";
									}else{
										value+=pv.getFullName()+"\n";
									}
	
								}
								//値
								excelTemplate.createCell(businessRow,cellNum++,Util.nullFilter(value));
							}
						}
					}
					/*** Compact Layout***/
					//コンパクトレイアウト
					excelTemplate.createTableHeaders(excelObjectSheet,"Compact Layout",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);//コンパクトレイアウト
					if(obj.getCompactLayouts().length>0){
						for( Integer i=0; i<obj.getCompactLayouts().length; i++){
							XSSFRow compactRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							CompactLayout ct =(CompactLayout)obj.getCompactLayouts()[i];
							cellNum = 1;
							Util.logger.debug("CompactLayout="+ct);
							//変更あり
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(compactRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("CompactLayout."+obj.getFullName()+"."+ct.getFullName()))));
							}
							//API名
							excelTemplate.createCell(compactRow,cellNum++,Util.nullFilter(ct.getFullName()));
							//ラベル
							excelTemplate.createCell(compactRow,cellNum++,Util.nullFilter(ct.getLabel()));
							String fields ="";
							for(String s: ct.getFields()){
								fields += this.apiToLabel(s) +"\n";
							}	
							//フィールド
							excelTemplate.createCell(compactRow,cellNum++,Util.nullFilter(fields));
	
						}
					}									
	
					/*** Custom Field ***/
					//创建Custom Field的Table(カスタム項目  リレーション)
					excelTemplate.createTableHeaders(excelObjectSheet,"Custom Field",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);//カスタム項目  リレーション
					Integer maxPicklistNum=0;
					Map<String,ValueSet> picklistMap = new HashMap<String,ValueSet>();
					Map<String,String> dependentMap = new HashMap<String,String>();
					Map<String,String> excelNameMap = new HashMap<String,String>();
					Map<String,LookupFilter> filterMap = new HashMap<String,LookupFilter>();
					Integer filterNum=0;
					if(obj.getFields().length>0){					
						for( Integer i=0; i<obj.getFields().length; i++ ){
							//Integer itemNum = excelObjectSheet.getLastRowNum()+1;
							//Create columnRow
							//XSSFRow fieldRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							//CustomField cf = (CustomField)obj.getFields()[i];
							
							CustomField cfall = (CustomField)obj.getFields()[i];
							Util.logger.debug("CustomField="+cfall);
							CustomField cf;
							if(cfall.getFullName().contains("__c")){
								cf = cfall;
								Integer itemNum = excelObjectSheet.getLastRowNum()+1;
								XSSFRow fieldRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								String fieldUpdateFlag=resultMap.get("CustomField."+obj.getFullName()+"."+cf.getFullName());
								String tmpFieldUpdateFlag=Util.getTranslate("IsChanged","NONE");
								cellNum=1;
								if(fieldUpdateFlag!=null){
									tmpFieldUpdateFlag=Util.getTranslate("IsChanged",fieldUpdateFlag);
								}
								//変更あり
								if(UtilConnectionInfc.modifiedFlag){
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(tmpFieldUpdateFlag));
								}
								//ラベル
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getLabel()));
								//API名
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getFullName()));
								if(cf.getFormula()!=null){
									//データ型
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("FieldType","FORMULA")+"("+Util.getTranslate("FieldType", Util.nullFilter(cf.getType()))+")");
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("FieldType", Util.nullFilter(cf.getType())));
								}
								//System.out.println("---------236-------cf.getFormula()="+cf.getFormula());
								//excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("FieldType", Util.nullFilter(cf.getType())));//データ型
								//大文字・小文字を区別する
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getCaseSensitive())));
								//デフォルト値
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getDefaultValue()));
								//説明
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getDescription()));
								//必須項目	
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getRequired())));
								//ユニーク	
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getUnique())));	
								//外部 ID		
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getExternalId())));					
								if(cf.getFormula()!=null){
									//数式
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getFormula()));
									String tmpFormulaTreatBlanksAs=Util.getTranslate("TreatBlanksAs", String.valueOf(""));
									if(cf.getFormulaTreatBlanksAs()!=null){
										tmpFormulaTreatBlanksAs=Util.getTranslate("TreatBlanksAs", String.valueOf(cf.getFormulaTreatBlanksAs()));
									}
									//空白項目の処理
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(tmpFormulaTreatBlanksAs));
									//System.out.println("-------------252--------tmpFormulaTreatBlanksAs="+tmpFormulaTreatBlanksAs);
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", Util.nullFilter("NotAvailable")));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", Util.nullFilter("NotAvailable")));
									//System.out.println("-------------256--------Util.getTranslate('Common', Util.nullFilter('NotAvailable'))="+Util.getTranslate("Common", Util.nullFilter("NotAvailable")));
								}		
								//ヘルプテキスト
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getInlineHelpText()));
								//文字数
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getLength()));
								if(String.valueOf(cf.getType())=="EncryptedText"){	
									//マスク文字
			                        excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("EncryptedFieldMaskChar",Util.nullFilter(cf.getMaskChar())));
			                        //マスク種別
			                        excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("EncryptedFieldMaskType",Util.nullFilter(cf.getMaskType())));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));								
								}
								if(String.valueOf(cf.getType()).contains("Picklist")){
									if(cf.getValueSet().getValueSetDefinition().getValue().length>0){
										maxPicklistNum=cf.getValueSet().getValueSetDefinition().getValue().length>maxPicklistNum?cf.getValueSet().getValueSetDefinition().getValue().length:maxPicklistNum;
										//Create columnRow
										Integer rowNo = itemNum;
										String plickliststring = cf.getFullName();
										String hyperVal = Util.makeNameValue(plickliststring);
										String displayVal =cf.getLabel();
										//参数 ： 当前sheet，行数，列数，hyperName,Cell表现值
										//選択リスト
										excelTemplate.createCellValue(excelObjectSheet,rowNo,cellNum++,hyperVal,displayVal);
										excelNameMap.put(cf.getFullName(), hyperVal);
										picklistMap.put(cf.getFullName(), cf.getValueSet());
										if(cf.getValueSet().getControllingField()!=null){
											dependentMap.put(cf.getValueSet().getControllingField(),cf.getFullName());
										}
									}								
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}
								List<String> fieldTypeWithScale = Arrays.asList("Currency","Number","Percent");
								if(fieldTypeWithScale.contains(String.valueOf(cf.getType()))){	
									//桁数
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getPrecision()-cf.getScale()));
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getScale()));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));								
								}
	
								List<String> fieldTypeWithLookup = Arrays.asList("Lookup","MasterDetail");
								if(fieldTypeWithLookup.contains(String.valueOf(cf.getType()))){		
									//関連先
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getReferenceTo()));
									//関連リストの表示ラベル
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getRelationshipLabel()));
									//子リレーション名
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getRelationshipName()));
									//主従関係順番
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getRelationshipOrder()));
									if(cf.getLookupFilter()!=null){
										filterNum++;
										//Create columnRow
										String filterString="LookupFilter"+filterNum;
										String hyperVal = Util.makeNameValue(filterString);
										String displayVal =filterString;
										//参数 ： 当前sheet，行数，列数，hyperName,Cell表现值
										//ルックアップ検索条件
										excelTemplate.createCellValue(excelObjectSheet,itemNum,cellNum++,hyperVal,displayVal);
										excelNameMap.put(cf.getFullName(), hyperVal);
										filterMap.put(cf.getFullName(), cf.getLookupFilter());
									
									}else{
										excelTemplate.createCell(fieldRow,cellNum++,"");
									}	
									String tmpDeleteConstraint=Util.getTranslate("DeleteConstraint", String.valueOf(""));
									if(cf.getDeleteConstraint()!=null){
										tmpDeleteConstraint=Util.getTranslate("DeleteConstraint", String.valueOf(cf.getDeleteConstraint()));
									}		
									//削除オプション	
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(tmpDeleteConstraint));
									//親の変更を許可
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getReparentableMasterDetail()));
									//共有設定(参照のみ)	
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getWriteRequiresMasterRead())));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}		
	
								if(String.valueOf(cf.getType())=="AutoNumber"){
									//表示形式		
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getDisplayFormat()));
									//開始番号
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getStartingNumber()));
									//既存レコードにも自動採番	
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getPopulateExistingRows()));						
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}
								if(String.valueOf(cf.getType())=="Html"){
									//マークアップを削除する
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getStripMarkup()));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}
								if(String.valueOf(cf.getType())=="Summary"){
									//集計する項目
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getSummarizedField()));
									//String filter=null;
									//for(FilterItem fi:cf.getSummaryFilterItems()){
									//	filter=fi.getValueField();							
									//}		
									//検索条件
									excelTemplate.createCell(fieldRow,cellNum++,ut.getFilterItem(obj.getFullName(), cf.getSummaryFilterItems()));
									//excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(filter));//検索条件
									//集計外部キー
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getSummaryForeignKey()));
									//積み上げ種別
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("summaryOperation", Util.nullFilter(cf.getSummaryOperation())));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}
								if(obj.getEnableFeeds()){
									//フィードの追跡
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getTrackFeedHistory())));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,"N/A");
								}
								if(obj.getEnableHistory()){
									//履歴の追跡
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getTrackHistory())));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,"N/A");
								}
								//excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getTrackFeedHistory())));//フィードの追跡
								//excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getTrackHistory())));//履歴の追跡
								//履歴トレンド
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getTrackTrending())));
	
								List<String> fieldTypeWithVisibleLines = Arrays.asList("LongTextArea","Html","MultiselectPicklist");
								if(fieldTypeWithVisibleLines.contains(String.valueOf(cf.getType()))){
									//表示行数
									excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getVisibleLines()));
								}else{
									excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("Common", "NotAvailable"));
								}
								//2015-6-8 added by duchuanchuan start
								//間接参照関係項目
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getReferenceTargetField()));
								//暗号化
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getEncrypted())));
								//外部データソースのテーブル列の名前
								excelTemplate.createCell(fieldRow,cellNum++,Util.nullFilter(cf.getExternalDeveloperName()));
								//isFilteringDisabled
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getIsFilteringDisabled())));
								//名前項目フラグ
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getIsNameField())));
								//並び替え可能フラグ
								excelTemplate.createCell(fieldRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(cf.getIsSortingDisabled())));
								//2015-6-8 added by duchuanchuan end
							}
						}
					}
				/***Lookup Filter added by duchuanchuan 2015/1/13**/
					//ルックアップ検索条件
					excelTemplate.createTableHeaders(excelObjectSheet,"Lookup Filter" ,excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(!filterMap.isEmpty()){
						int no = 1;
						for(String s:filterMap.keySet()){
							excelTemplate.createCellName(excelNameMap.get(s),objectName,excelObjectSheet.getLastRowNum()+2);
							XSSFRow filterRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							LookupFilter lf = filterMap.get(s);
							cellNum = 1;
							//No.
							excelTemplate.createCell(filterRow, cellNum++, "lookupFilter_"+no++);
							//有効
							excelTemplate.createCell(filterRow, cellNum++, Util.getTranslate("BooleanValue",Util.nullFilter(lf.getActive())));
							//検索条件ロジック
							excelTemplate.createCell(filterRow, cellNum++,Util.nullFilter(lf.getBooleanFilter()));
							//説明
							excelTemplate.createCell(filterRow, cellNum++, Util.nullFilter(lf.getDescription()));
							//ルックアップウィンドウテキスト
							excelTemplate.createCell(filterRow, cellNum++, Util.nullFilter(lf.getInfoMessage()));
							/*FilterItem[] fitems = lf.getFilterItems();
							String ftstr = "";
							for(FilterItem ft:fitems){
								ftstr += ft.getField()+" ";
								ftstr += Util.getTranslate("FilterOperation",String.valueOf(ft.getOperation()))+" ";
								ftstr += ft.getValue()+"(";
								ftstr += String.valueOf(ft.getValueField())+")\n";							
							}
							excelTemplate.createCell(filterRow, cellNum++,Util.nullFilter(ftstr));//検索条件
							*/
							//検索条件
							excelTemplate.createCell(filterRow,cellNum++,ut.getFilterItem(obj.getFullName(),lf.getFilterItems()));
							//エラーメッセージ
							if(lf.getErrorMessage()==null){								
								excelTemplate.createCell(filterRow, cellNum++, Util.getTranslate("ERRORMESSAGE","DEFAULT"));
							}else{
								excelTemplate.createCell(filterRow, cellNum++, Util.nullFilter(lf.getErrorMessage()));
							}
							//条件種別
							excelTemplate.createCell(filterRow, cellNum++, Util.getTranslate("BooleanValue",Util.nullFilter(lf.getIsOptional())));
						}
					}
					
					/************end by duchuanchuan******/
					/*** List Views***/
					//リストビュー
					excelTemplate.createTableHeaders(excelObjectSheet,"List Views",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getListViews().length>0){						
						for( Integer i=0; i<obj.getListViews().length; i++){
							XSSFRow listRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							ListView lv = (ListView)obj.getListViews()[i];
							cellNum = 1;
							Util.logger.debug("ListView="+lv);
							//変更あり
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(listRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("ListView."+obj.getFullName()+"."+lv.getFullName()))));
							}
							//API名
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getFullName()));
							//ビュー名
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getLabel()));
							//検索条件ロジック
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getBooleanFilter()));
							//ディビジョン
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getDivision()));
							//フィルター範囲
							excelTemplate.createCell(listRow,cellNum++,Util.getTranslate("FilterScope", Util.nullFilter(lv.getFilterScope())));
							String item="";
							for(ListViewFilter fi:lv.getFilters()){
								item+=fi.getField()+" ";
								item+=fi.getOperation()+" ";
								item+=fi.getValue()+"\n";
							}
							//フィルター	
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(item));
							//言語
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getLanguage()));
							//キュー
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(lv.getQueue()));	
							Util.logger.debug(lv.getSharedTo());
							//共有先		
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(ut.getSharedTo(lv.getSharedTo())));				
						}
					}						
	
					/*** Record Type***/
					Map<String,List<String>> recordpicklistMap = new TreeMap<String,List<String>>();
					//レコードタイプ
					excelTemplate.createTableHeaders(excelObjectSheet,"Record Type",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getRecordTypes().length>0){
						for( Integer i=0; i<obj.getRecordTypes().length; i++){
							XSSFRow listRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							RecordType rt = (RecordType)obj.getRecordTypes()[i];	
							cellNum = 1;
							Util.logger.debug("ListView="+rt);
							//変更あり
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(listRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("RecordType."+obj.getFullName()+"."+rt.getFullName()))));
							}
							//表示ラベル
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(rt.getLabel()));
							//API名
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(rt.getFullName()));
							//説明
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(rt.getDescription()));
							//有効
							excelTemplate.createCell(listRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(rt.getActive())));
							//ビジネスプロセス
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(rt.getBusinessProcess()));
							//コンパクトレイアウトの割り当て
							excelTemplate.createCell(listRow,cellNum++,Util.nullFilter(rt.getCompactLayoutAssignment()));
							if(rt.getPicklistValues().length>0){															
								for(int j=0;j<rt.getPicklistValues().length;j++){
									Integer rownum;
									if(j==0){
										rownum = listRow.getRowNum();
									}else{	
										XSSFRow srow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
										//for cell style
										for(int n=1;n<cellNum;n++){
											excelTemplate.createCell(srow,n,"");
										}
										rownum=srow.getRowNum();
									}
									RecordTypePicklistValue rp = rt.getPicklistValues()[j];
									
									String hyperVal =  Util.makeNameValue(rt.getFullName()+"_"+rp.getPicklist());									
									List<String> l = new ArrayList<String>();
									List<String> inMap = new ArrayList<String>();
									for(PicklistValue pv:rp.getValues()){										
										l.add(pv.getFullName());
									}
									if(picklistMap.containsKey(rp.getPicklist())){
									    for(CustomValue pv:picklistMap.get(rp.getPicklist()).getValueSetDefinition().getValue()){								    	
									    	if(l.contains(pv.getFullName())){
									    		inMap.add(pv.getFullName());
									    	}
									    }
									}
								    if(inMap != null){
								    	recordpicklistMap.put(hyperVal, inMap);
								    }
								    //選択リスト
									excelTemplate.createCellValue(excelObjectSheet,rownum,cellNum,hyperVal,rp.getPicklist());
								}
								cellNum++;
							}
						}
					}					
	
					/*** Referenced PickLists ***/
					//创建Referenced PickLists的Table(選択リスト詳細参照)
					excelTemplate.createTableHeaders(excelObjectSheet,"Referenced PickLists",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					for( cellNum=1;cellNum<=maxPicklistNum;cellNum++){
						XSSFRow pickHeaderRows =  excelObjectSheet.getRow(excelObjectSheet.getLastRowNum());
						XSSFCell cell = pickHeaderRows.createCell(cellNum+1);
						cell.setCellValue(Util.getTranslate("CUSTOMFIELD", "PICKLISTVALUE")+cellNum);
						cell.setCellStyle(excelTemplate.createCHeaderStyle());
					}
					for(String s:picklistMap.keySet()){
						excelTemplate.createCellName(excelNameMap.get(s),objectName,excelObjectSheet.getLastRowNum()+2);
						XSSFRow pickRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						cellNum=1;
						excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(s));
						for(int j=0;j<picklistMap.get(s).getValueSetDefinition().getValue().length;j++){
							String color="";
							if(picklistMap.get(s).getValueSetDefinition().getValue()[j].getColor()!=null){
								color="("+picklistMap.get(s).getValueSetDefinition().getValue()[j].getColor()+")";								
							}
							excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(picklistMap.get(s).getValueSetDefinition().getValue()[j].getFullName()+color));							
						}
						for(int k=picklistMap.get(s).getValueSetDefinition().getValue().length;k<maxPicklistNum;k++){
							excelTemplate.createCell(pickRow,cellNum++,"");
						}
					}
					cellNum=1;
					if(recordpicklistMap != null){
						for(String s:recordpicklistMap.keySet()){
							excelTemplate.createCellName(s,objectName,excelObjectSheet.getLastRowNum()+2);
							XSSFRow pickRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(s.substring(5)));
							if(recordpicklistMap.containsKey(s)){
								for(int j=0;j<recordpicklistMap.get(s).size();j++){
									excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(recordpicklistMap.get(s).get(j)));							
								}
							}
							//do empty item cell
							if(recordpicklistMap.containsKey(s)){
								for(int k=recordpicklistMap.get(s).size();k<maxPicklistNum;k++){
									excelTemplate.createCell(pickRow,cellNum++,"");
								}	
							}
							cellNum=1;
						}
					}
	
	
					/*** Search　Layout ***/
					//创建Search　Layout的Table(検索レイアウト)
					excelTemplate.createTableHeaders(excelObjectSheet,"Search　Layout",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getSearchLayouts()!=null){
						SearchLayouts sl = obj.getSearchLayouts();
						if(sl.getCustomTabListAdditionalFields().length>0||
								sl.getExcludedStandardButtons().length>0||
								sl.getListViewButtons().length>0||
								sl.getLookupDialogsAdditionalFields().length>0||
								sl.getLookupFilterFields().length>0||
								sl.getLookupPhoneDialogsAdditionalFields().length>0||
								sl.getSearchFilterFields().length>0||
								sl.getSearchResultsAdditionalFields().length>0||
								sl.getSearchResultsCustomButtons().length>0){
							XSSFRow searchRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);							
							cellNum = 1;
							Util.logger.debug("search layouts="+sl);
							String customTabListAdditionalFields="";
							for(String s:sl.getCustomTabListAdditionalFields()){
								customTabListAdditionalFields+=this.apiToLabel(s)+"\n";
							}
							//タブ
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(customTabListAdditionalFields));
							String excludedStandardButtons="";
							for(String s:sl.getExcludedStandardButtons()){
								excludedStandardButtons+=s+"\n";
							}
							//除外する標準ボタン
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(excludedStandardButtons));
							String listViewButtons="";
							for(String s:sl.getListViewButtons()){
								listViewButtons+=s+"\n";
							}
							//リストビューボタン
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(listViewButtons));
							String lookupDialogsAdditionalFields="";
							for(String s:sl.getLookupDialogsAdditionalFields()){
								lookupDialogsAdditionalFields+=this.apiToLabel(s)+"\n";
							}
							//ルックアップダイアログ項目
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(lookupDialogsAdditionalFields));
							String lookupFilterFields="";
							for(String s:sl.getLookupFilterFields()){
								lookupFilterFields+=this.apiToLabel(s)+"\n";
							}
							//ルックアップ検索条件項目
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(lookupFilterFields));
							String lookupPhoneDialogsAdditionalFields="";
							for(String s:sl.getLookupPhoneDialogsAdditionalFields()){
								lookupPhoneDialogsAdditionalFields+=this.apiToLabel(s)+"\n";
							}
							//ルックアップ電話ダイアログ項目
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(lookupPhoneDialogsAdditionalFields));
							String searchFilterFields="";
							for(String s:sl.getSearchFilterFields()){
								searchFilterFields+=this.apiToLabel(s)+"\n";
							}
							//検索条件項目
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(searchFilterFields));
							String searchResultsAdditionalFields="";
							for(String s:sl.getSearchResultsAdditionalFields()){
								searchResultsAdditionalFields+=this.apiToLabel(s)+"\n";
							}
							//検索結果項目
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(searchResultsAdditionalFields));
							String searchResultsCustomButtons="";
							for(String s:sl.getSearchResultsCustomButtons()){
								searchResultsCustomButtons+=s+"\n";
							}
							//検索結果ボタン
							excelTemplate.createCell(searchRow,cellNum++,Util.nullFilter(searchResultsCustomButtons));
						}
												
					}
	
					/*** Web Link***/
					//创建Web Link的Table(Webリンク)
					excelTemplate.createTableHeaders(excelObjectSheet,"Web Link",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getWebLinks().length>0){
						for(int i=0;i<obj.getWebLinks().length;i++){
							XSSFRow webRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							WebLink wl = obj.getWebLinks()[i];
							cellNum=1;
							Util.logger.debug("WebLink="+wl);
							//変更あり
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("WebLink."+obj.getFullName()+"."+wl.getFullName()))));
							}
							//API名
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getFullName()));
							//リンクタイプ
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("WebLinkType", Util.nullFilter(wl.getLinkType())));		
							//表示ラベル
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getMasterLabel()));
							//表示の種類
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("WebLinkDisplayType", Util.nullFilter(wl.getDisplayType())));
							//使用場所
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getAvailability()));
							//説明
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getDescription()));
							//エンコード
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getEncodingKey()));
							//メニューバーの表示
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getHasMenubar())));
							//スクロールバーの表示
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getHasScrollbars())));
							//ツールバーの表示
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getHasToolbar())));
							//高さ (ピクセル)
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getHeight()));
							//幅 (ピクセル)
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getWidth()));
							//サイズ変更可能
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getIsResizable())));
							//動作
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BEHAVIOR",Util.nullFilter(wl.getOpenType())));
							//ウィンドウの位置
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("WINDOWPOSITION",Util.nullFilter(wl.getPosition())));
							//保護
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getProtected())));
							//複数レコード選択
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getRequireRowSelection())));
							//Sコントロール
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getScontrol()));
							//アドレスバーの表示
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getShowsLocation())));
							//ステータスバーの表示
							excelTemplate.createCell(webRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wl.getShowsStatus())));
							//URL
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getUrl()));
							//Visualforceページ
							excelTemplate.createCell(webRow,cellNum++,Util.nullFilter(wl.getPage()));
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
	
					/*** Sharing Reason***/
					//创建Sharing Reason的Table(Apex 共有の理由)
					excelTemplate.createTableHeaders(excelObjectSheet,"Sharing Reason",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getSharingReasons().length>0){
						for(int i=0;i<obj.getValidationRules().length;i++){
							XSSFRow sharingRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							SharingReason sr = obj.getSharingReasons()[i];
							cellNum=1;
							//変更あり
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(sharingRow,cellNum++,Util.getTranslate("IsChanged",Util.nullFilter(resultMap.get("SharingReason."+obj.getFullName()+"."+sr.getFullName()))));
							}
							//理由名
							excelTemplate.createCell(sharingRow,cellNum++,Util.nullFilter(sr.getFullName()));
							//表示ラベル
							excelTemplate.createCell(sharingRow,cellNum++,Util.nullFilter(sr.getLabel()));
						}
					}
	
					/*** Apex Sharing Recalculation***/
					//创建Apex Sharing Recalculation的Table(Apex共有再適用)
					excelTemplate.createTableHeaders(excelObjectSheet,"Apex Sharing Recalculation",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getSharingRecalculations().length>0){
						for(int i=0;i<obj.getSharingRecalculations().length;i++){
							XSSFRow apexRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							cellNum=1;
							SharingRecalculation sr = obj.getSharingRecalculations()[i];
							//クラス名
							excelTemplate.createCell(apexRow,cellNum++,Util.nullFilter(sr.getClassName()));
	
						}
					}					
	
					/*** Dependent Picklist ***/
					//创建Referenced PickLists的Table(項目の連動関係)
					excelTemplate.createTableHeaders(excelObjectSheet,"Dependent Picklist",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					if(!dependentMap.isEmpty()){
						for(String s:dependentMap.keySet()){
							XSSFRow headRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							headRow.createCell(1).setCellValue(apiToLabel(s)+" -- "+apiToLabel(dependentMap.get(s)));
							cellNum=1;
							XSSFRow controlledRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
							excelTemplate.createCell(controlledRow,cellNum++,"");
							//modify by vp cheng 15-9-10 start
							/*for(int x=0;x<picklistMap.get(dependentMap.get(s)).getPicklistValues().length;x++){
								excelTemplate.createCell(controlledRow,cellNum++,Util.nullFilter(picklistMap.get(dependentMap.get(s)).getPicklistValues()[x].getFullName()));	
								//System.out.println("-------------picklistMap.get(dependentMap.get(s)).getPicklistValues()[x].getFullName()="+picklistMap.get(dependentMap.get(s)).getPicklistValues()[x].getFullName());
							}*/
							if (picklistMap.containsKey(s)) {
								for(int x=0;x<picklistMap.get(s).getValueSetDefinition().getValue().length;x++){
									excelTemplate.createCell(controlledRow,cellNum++,Util.nullFilter(picklistMap.get(s).getValueSetDefinition().getValue()[x].getFullName()));		
								}
							}	
							cellNum=1;
							if (picklistMap.containsKey(dependentMap.get(s))) {
								Util.logger.debug("What the total value here : "+picklistMap.get(dependentMap.get(s)).getValueSetDefinition().getValue().length);
								for(int j=0;j<picklistMap.get(dependentMap.get(s)).getValueSetDefinition().getValue().length;j++){
									Util.logger.debug("What the name here : "+picklistMap.get(dependentMap.get(s)).getValueSetDefinition().getValue()[j].getFullName());
									XSSFRow pickRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
									//pickRow.createCell(1).setCellValue(picklistMap.get(s).getPicklistValues()[j].getFullName());
									excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(picklistMap.get(dependentMap.get(s)).getValueSetDefinition().getValue()[j].getFullName()));
									if (picklistMap.containsKey(s)) {
										Map<String,String[]> valueSettingMap = new HashMap<String,String[]>();
										if(picklistMap.get(dependentMap.get(s)).getValueSettings() != null){
											for(ValueSettings vs:  picklistMap.get(dependentMap.get(s)).getValueSettings()){
												valueSettingMap.put(vs.getValueName(), vs.getControllingFieldValue());
											}
										}
										for(int k=0;k<picklistMap.get(s).getValueSetDefinition().getValue().length;k++){
											if(isIn(picklistMap.get(s).getValueSetDefinition().getValue()[k].getFullName(),valueSettingMap,picklistMap.get(dependentMap.get(s)).getValueSetDefinition().getValue()[j].getFullName())){							
												excelTemplate.createCell(pickRow,cellNum++,Util.getTranslate("BooleanValue", "TRUE"));
											}else{
												excelTemplate.createCell(pickRow,cellNum++,Util.getTranslate("BooleanValue", "FALSE"));
											}
										}
									}
									
									
									
									cellNum=1;
								}
							}
							/*if (picklistMap.containsKey(s)) {
								for(int j=0;j<picklistMap.get(s).getPicklistValues().length;j++){
									
									XSSFRow pickRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
									//pickRow.createCell(1).setCellValue(picklistMap.get(s).getPicklistValues()[j].getFullName());
									excelTemplate.createCell(pickRow,cellNum++,Util.nullFilter(picklistMap.get(s).getPicklistValues()[j].getFullName()));
									//System.out.println("-------------picklistMap.get(s).getPicklistValues()[j].getFullName()="+picklistMap.get(s).getPicklistValues()[j].getFullName());
									for(int k=0;k<picklistMap.get(dependentMap.get(s)).getPicklistValues().length;k++){
										if(isIn(picklistMap.get(s).getPicklistValues()[j].getFullName(),picklistMap.get(dependentMap.get(s)).getPicklistValues()[k].getControllingFieldValues())){							
											excelTemplate.createCell(pickRow,cellNum++,Util.getTranslate("BooleanValue", "TRUE"));
										}else{
											excelTemplate.createCell(pickRow,cellNum++,Util.getTranslate("BooleanValue", "FALSE"));
										}
									}
									cellNum=1;
								}
							}*/
							//modify by vp cheng 15-9-10 end
						}
					}
					excelTemplate.adjustColumnWidth(excelObjectSheet);					
				}else {
					Util.logger.warn("Empty metadata.");
				}	
				//
				if(util.createExcel(workBook, excelTemplate, type, objectsList.size(), lastIndex)){
					excelTemplate.CreateWorkBook(type);
					workBook = excelTemplate.workBook;
				}
			}
		}catch (Exception e) {
			Util.logger.error("ReadCustomObject failure.");
			Util.logger.error("",e);
		}
		Util.logger.info("ReadCustomObject End.");
	}
	
	public String apiToLabel(String apiName){
		String result="";
		if(apiToLabelMap.get(apiName)!=null){
			result=apiToLabelMap.get(apiName);
		}else{
			result=apiName;
		}
		return result;
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
	/**
	 * JAVA判断字符串数组中是否包含某字符串元素
	 *
	 * @param substring 某字符串
	 * @param source 源字符串数组
	 * @return 包含则返回true，否则返回false
	 */
	public static boolean isIn(String substring, Map<String,String[]> stringMap, String dependentName) {
		if(stringMap == null){
			return false;
		}
		if(stringMap.containsKey(dependentName)){
			for(int i =0; i < stringMap.get(dependentName).length; i++){
				String checkName = stringMap.get(dependentName)[i];
				if(checkName.equals(substring)){
					return true;
				}
			}
		}
		/*if (source == null || source.length == 0) {
			return false;
		}
		for (int i = 0; i < source.length; i++) {
			String aSource = source[i];
			if (aSource.equals(substring)) {
				return true;
			}
		}*/
		return false;
	}
}
