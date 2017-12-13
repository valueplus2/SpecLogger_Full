package source;

import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.CustomFieldTranslation;
import com.sforce.soap.metadata.CustomObjectTranslation;
import com.sforce.soap.metadata.LayoutSectionTranslation;
import com.sforce.soap.metadata.LayoutTranslation;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.ObjectNameCaseValue;
import com.sforce.soap.metadata.PicklistValueTranslation;
import com.sforce.soap.metadata.QuickActionTranslation;
import com.sforce.soap.metadata.RecordTypeTranslation;
import com.sforce.soap.metadata.SharingReasonTranslation;
import com.sforce.soap.metadata.ValidationRuleTranslation;
import com.sforce.soap.metadata.WebLinkTranslation;
import com.sforce.soap.metadata.WorkflowTaskTranslation;
import com.sforce.ws.ConnectionException;

public class ReadCustomObjectTranslationSync {
	private XSSFWorkbook workbook;
   public void readCustomObjectTranslation(String type,List<String> objectsList) throws Exception{
	   Util.logger.info("readCustomObjectTranslation Start.");
	   Util ut = new Util();
	   Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workbook = excelTemplate.workBook;
		//Create Catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		if(mdInfos.size()>0){
			for(Metadata md:mdInfos){//have eight
				CustomObjectTranslation cot = (CustomObjectTranslation)md;
				String sheetname=Util.makeSheetName(cot.getFullName());
				XSSFSheet customObjectTransSheet = excelTemplate.createSheet(Util.cutSheetName(sheetname));
				//objectNameCaseValue list
				Map<String,ObjectNameCaseValue> objNameCaseValuesMap  = new HashMap<String,ObjectNameCaseValue>(); 
				Map<String[],PicklistValueTranslation[]> picklistValueMap = new HashMap<String[],PicklistValueTranslation[]>();
				//创建目录信息
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,customObjectTransSheet,Util.cutSheetName(sheetname),sheetname);
				/**********CustomObjectTranslation************/
				//create CustmObjectTranslation table
				//Integer itemNum = customObjectTransSheet.getLastRowNum()+1;
				//カスタムオブジェクト翻訳
				excelTemplate.createTableHeaders(customObjectTransSheet, "Custom Object Translation",customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum);
				int cellNum=1;
				XSSFRow columnRow = customObjectTransSheet.createRow(customObjectTransSheet.getLastRowNum()+1);
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"customObjectTransSheet."+cot.getFullName()));
					
				}
				//API参照名_翻訳言語
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cot.getFullName()));
				//言語
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cot.getGender()==null?"":Util.nullFilter(cot.getGender())));
				//レコード名
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cot.getNameFieldLabel()));
				//母音で始まる場合はチェック
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cot.getStartsWith()==null?"":Util.getTranslate("startsWith", Util.nullFilter(cot.getStartsWith()))));
				
			    if(cot.getCaseValues().length>0){
			    	//maxCaseValuesNum = cot.getCaseValues().length;
			    	
			    Integer rowNum = customObjectTransSheet.getLastRowNum();
                  for(ObjectNameCaseValue cs:cot.getCaseValues()){
                	  
  			    	  String hyperVal = Util.makeNameValue(cot.getFullName());
                      String displayVal = cot.getFullName();
                      objNameCaseValuesMap.put(hyperVal,cs );
                      //参数 ： 当前sheet，行数，列数，hyperName,Cell表现值
                      //文法
                      excelTemplate.createCellValue(customObjectTransSheet, rowNum,cellNum, hyperVal, displayVal);
			    	}
			    }else{
			    	excelTemplate.createCell(columnRow,cellNum,"");
			    }
			    cellNum++;
				
				
 				/***************Custom Field Translation***********/
				//create table custonFieldTranslation
				Integer lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
				//カスタム項目翻訳
				excelTemplate.createTableHeaders(customObjectTransSheet, "Custom Field Translation",lastRow);
				CustomFieldTranslation[] cfts = cot.getFields();
				if(cfts.length>0){
					cellNum=1;
					//Integer curRow = lastRow;
					for(CustomFieldTranslation cft:cfts){
						lastRow = customObjectTransSheet.getLastRowNum()+1;
						XSSFRow newRow = customObjectTransSheet.createRow(lastRow);
						//API参照名
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getName()));
						//newRow.createCell(0).setCellValue(resultMap.get("CustomFieldTranslation."+cot.getFullName()+"."+cft.getName()));
						if(cft.getCaseValues().length>0){
							
							Integer cvRowNum = lastRow;
							for(ObjectNameCaseValue cs:cft.getCaseValues()){
								String hyperVal = Util.makeNameValue(cot.getFullName());
								
								String displayVal = cft.getName();
								//文法
								excelTemplate.createCellValue(customObjectTransSheet, cvRowNum++, cellNum, hyperVal, displayVal);
								objNameCaseValuesMap.put(hyperVal, cs);
							}
							cellNum++;
						}else{
							excelTemplate.createCell(newRow,cellNum++,"");
						}
						//言語
						excelTemplate.createCell(newRow,cellNum++,cft.getGender()==null?"":Util.nullFilter(cft.getGender()));
						//ヘルプテキストの翻訳
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getHelp()));
						//項目表示ラベルの翻訳
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getLabel()));
						//cfts[i].getCaseValues()[0
						if(cft.getLookupFilter()!=null){
							//エラーメッセージの翻訳
							excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getLookupFilter().getErrorMessage()));
							//情報メッセージの翻訳
							excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getLookupFilter().getInformationalMessage()));
						}else{
							excelTemplate.createCell(newRow,cellNum++,"");
							excelTemplate.createCell(newRow,cellNum++,"");
						}
						
	                    
	                    PicklistValueTranslation[] pvs = cft.getPicklistValues();
	                    
	                    if(pvs.length>0){     	                    	
                    		String displayVal =cft.getName();                 	        
	                    	String hyperVal = Util.makeNameValue(cft.getName());
	                    	//選択リスト値
	                    	excelTemplate.createCellValue(customObjectTransSheet, newRow.getRowNum(), cellNum++, hyperVal, displayVal);
                    		String[] temp = new String[]{hyperVal,displayVal};
	                    	picklistValueMap.put(temp, pvs);
	                    }else{
	                    	excelTemplate.createCell(newRow,cellNum++,"");
	                    }
	                    //関連リストの表示ラベルの翻訳
	                    excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cft.getRelationshipLabel()));
	                    //母音で始まる場合はチェック
	                    excelTemplate.createCell(newRow,cellNum++,cft.getStartsWith()==null?"":Util.getTranslate("startsWith",Util.nullFilter(cft.getStartsWith())));
	                    cellNum=1;
	                 }
				}
				/*****************ObjectNameCaseValue****************/
				//create table custom object case value
				lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
				//オブジェクトの文法
				excelTemplate.createTableHeaders(customObjectTransSheet, "Object Name Case Value",lastRow);
				//ObjectNameCaseValue[] cocvs = cot.getCaseValues();
				if(!objNameCaseValuesMap.isEmpty()){
					for(String s:objNameCaseValuesMap.keySet()){
						//create a new row
						Integer num = customObjectTransSheet.getLastRowNum()+1;
						Util.logger.debug("ssssv++++="+s);
						XSSFRow newRow = customObjectTransSheet.createRow(num);
						excelTemplate.createCellName(s, sheetname, num+1);
						//cocv.getCaseType();
						cellNum=1;
						//API参照名
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(s));
						ObjectNameCaseValue cocv = objNameCaseValuesMap.get(s);
						//定冠詞／不定冠詞
						excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("article",Util.nullFilter(cocv.getArticle()==null?"None":cocv.getArticle())));
						//格
						excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("caseType", Util.nullFilter(cocv.getCaseType()==null?"None":cocv.getCaseType())));
						//複数形
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cocv.getPlural()));
						//所有格
						excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("possessive", Util.nullFilter(cocv.getPossessive()==null?"None":cocv.getPossessive())));
						//表示ラベル
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cocv.getValue()));
					    
				    }
				}
				
				/**********CustomLayout Translation********/
			    //create tale layout translation
				lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
				//レイアウトセッションの翻訳
				excelTemplate.createTableHeaders(customObjectTransSheet, "Layout Translation", lastRow);
				LayoutTranslation[] lts = cot.getLayouts();
				if(lts.length>0){
					lastRow = customObjectTransSheet.getLastRowNum()+1;
					
					for(LayoutTranslation lt:lts){
						//create row
						cellNum=1;
						XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
						//newRow.createCell(0).setCellValue(resultMap.get("LayoutTranslation."+cot.getFullName()+"."+lt.getLayout()));
						//レイアウト名
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(lt.getLayout()));
						//レイアウトタイプ
						excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(lt.getLayoutType()));
					    String sections = "";
					    LayoutSectionTranslation[] lsts = lt.getSections();
					    for(LayoutSectionTranslation lst:lsts){
					    	sections = sections+lst.getLabel()+":"+lst.getSection()+"\r\n";
					    }
					    //セッション名
					    excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(sections));
					}
				}
				/****************PicklistValueTranslation*************/
				Integer num = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
				//選択リスト値の翻訳
				excelTemplate.createTableHeaders(customObjectTransSheet, "Picklist Value Translation", num);
				if(!picklistValueMap.isEmpty()){
					String temName="";
					for(String[] s:picklistValueMap.keySet()){
						PicklistValueTranslation[] pvs = picklistValueMap.get(s);
						
						for(PicklistValueTranslation pv:pvs){
							Integer pvRowNum = customObjectTransSheet.getLastRowNum()+1;
							cellNum=1;
							XSSFRow newRow = customObjectTransSheet.createRow(pvRowNum);
							if(temName!=s[1]){								
								excelTemplate.createCellName(s[0], sheetname, pvRowNum+1);
								//項目API参照名
								excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(s[1]));
								temName=s[1];
							}else{
								excelTemplate.createCell(newRow,cellNum++,"");
							}
							//翻訳言語API参照名
							excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(cot.getFullName()));
							//マスタ選択リスト値の表示ラベル
							excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(pv.getMasterLabel()));
							//選択リスト値の表示ラベルの翻訳
							excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(pv.getTranslation()));
						}
					}
				}
			   
				/*****************QuickActionTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //アクションの翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Quick Action Translation", lastRow);
			   QuickActionTranslation[] qats = cot.getQuickActions();
			   if(qats.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum()+1;
				   for(QuickActionTranslation qat:qats){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
					   // newRow.createCell(0).setCellValue(resultMap.get("QuickActionTranslation."+cot.getFullName()));
					   //アクション
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(qat.getName()));
					   //表示ラベル
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(qat.getLabel()));
				   } 
			   }
			   /*****************RecordTypeTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //レコードタイプの翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Record Type Translation", lastRow);
			   RecordTypeTranslation[] rtts = cot.getRecordTypes();
			   if(rtts.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum();
				   for(RecordTypeTranslation rtt:rtts){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow+1);
					   //レコードタイプ
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rtt.getName()));
				       //表示ラベル
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(rtt.getLabel()));
				   } 
			   }
			   /****************SharingReasonTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //Apex共有の理由の翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Sharing Reason Translation", lastRow);
			   SharingReasonTranslation[] srts = cot.getSharingReasons();
			   if(srts.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum()+1;
				   for(SharingReasonTranslation srt:srts){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
					   // newRow.createCell(0).setCellValue(resultMap.get("SharingReasonTranslation."+cot.getFullName()+"."+srt.getName()));
					   //Apex共有の理由
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(srt.getName()));
					   //表示ラベル
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(srt.getLabel()));
				   } 
			   }
			   /*****************ValidationRuleTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //入力エラーメッセージの翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Validation Rule Translation", lastRow);
			   ValidationRuleTranslation[] vrts = cot.getValidationRules();
			   if(vrts.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum()+1;
				   for(ValidationRuleTranslation vrt:vrts){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
					   //newRow.createCell(0).setCellValue(resultMap.get("ValidationRuleTranslation."+cot.getFullName()+"."+vrt.getName()));
					   //入力規則
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(vrt.getName()));
					   //エラーメッセージ
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(vrt.getErrorMessage()));
				   } 
			   }
			   
			   /*****************WebLinkTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //ボタンとリンクの表示ラベルの翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Web Link Translation", lastRow);
			   WebLinkTranslation[] wlts = cot.getWebLinks();
			   if(wlts.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum()+1;
				   for(WebLinkTranslation wlt:wlts){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
					   //newRow.createCell(0).setCellValue(resultMap.get("WebLinkTranslation."+cot.getFullName()+"."+wlt.getName()));
					   //Webタブ
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(wlt.getName()));
					   //表示ラベル
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(wlt.getLabel()));
				   } 
			   }
			   
			   /*****************WorkflowTaskTranslation*****************/
			   lastRow = customObjectTransSheet.getLastRowNum()+Util.RowIntervalNum;
			   //ワークフローToDoの翻訳
			   excelTemplate.createTableHeaders(customObjectTransSheet, "Workflow Task Translation", lastRow);
			   WorkflowTaskTranslation[] wtts = cot.getWorkflowTasks();
			   if(wtts.length>0){
				   lastRow = customObjectTransSheet.getLastRowNum()+1;
				   for(WorkflowTaskTranslation wtt:wtts){
					   //create row
					   cellNum=1;
					   XSSFRow newRow = customObjectTransSheet.createRow(lastRow++);
					   // newRow.createCell(0).setCellValue(resultMap.get("WorkflowTaskTranslation."+cot.getFullName()+"."+wtt.getName()));
					   //ワークフローToDo
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(wtt.getName()));
					   //コメント
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(wtt.getDescription()));
					   //件名
					   excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(wtt.getSubject()));
				   } 
			   }
			   excelTemplate.adjustColumnWidth(customObjectTransSheet);
			}
			
		}else{
			Util.logger.info("empty metadata");
		}
		if(workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("***no result to export!!!");
		}
		Util.logger.info("readCustomObjectTranslation End.");
   }
   /*s*/
}
