package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.NetworkBranding;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.NavigationLinkSet;
import com.sforce.soap.metadata.NavigationMenuItem;
import com.sforce.soap.metadata.NavigationSubMenu;
import com.sforce.soap.metadata.Network;
import com.sforce.soap.metadata.NetworkMemberGroup;
import com.sforce.soap.metadata.NetworkPageOverride;
import com.sforce.soap.metadata.NetworkPageOverrideSetting;
import com.sforce.soap.metadata.NetworkStatus;
import com.sforce.soap.metadata.NetworkTabSet;
import com.sforce.soap.metadata.ReputationBranding;
import com.sforce.soap.metadata.ReputationLevel;
import com.sforce.soap.metadata.ReputationLevelDefinitions;
import com.sforce.soap.metadata.ReputationPointsRule;
import com.sforce.soap.metadata.ReputationPointsRules;
import com.sforce.ws.ConnectionException;

public class ReadNetworkSync {
	private XSSFWorkbook workbook;
	
	public void readNetwork(String type, List<String> objectsList,List<String> objectsList1)
			throws Exception {
		Util.logger.info("ReadNetworkSync Started.");
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;				
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		System.out.println("objectsList---------"+objectsList);
		System.out.println("mdInfos---------"+mdInfos);
		//add by cheng 2017-9-12 start
		System.out.println("objectsList1---------"+objectsList1);
		Map<String,NetworkBranding> netWorkBrandMap = new HashMap<String,NetworkBranding>();
		if(objectsList1!=null){
			List<Metadata> mdInfos1 = ut.readMateData("NetworkBranding", objectsList1);
			System.out.println("objectsList1---------"+objectsList);
			System.out.println("mdInfos1---------"+mdInfos1);
			
			for(int i = 0; i < mdInfos1.size(); i++){
				NetworkBranding nwb =  (NetworkBranding)mdInfos1.get(i);
				System.out.println("nwb---------"+nwb);
				netWorkBrandMap.put(nwb.getNetwork(),nwb);
			}
		}
		//add by cheng 2017-9-12 end
		
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		//XSSFSheet catalog = excelTemplate.createCatalogSheet();
		this.workbook = excelTemplate.workBook;
		
		// loop network detail
		for (int i = 0; i < mdInfos.size(); i++) {
			Network network = (Network) mdInfos.get(i);
			Util.logger.info("network.getFullName()="+network.getFullName());
			String sheetName = Util.makeSheetName(network.getFullName());
			Util.logger.info("network.getFullName().SheetName="+sheetName);
			sheetName = URLDecoder.decode(sheetName, "UTF-8");
			XSSFSheet networkSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
			excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, networkSheet,networkSheet.getSheetName(),sheetName);
			
			//Create network Table(ネットワーク)
			excelTemplate.createTableHeaders(networkSheet, "Network",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			XSSFRow rowNetwork = networkSheet.createRow(networkSheet.getLastRowNum() + 1);
			int cellNum=1;
					
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getAllowedExtensions()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getAllowInternalUserLogin())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getAllowMembersToFlag())));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getCaseCommentEmailTemplate()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getChangePasswordTemplate()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getDescription()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getEmailSenderAddress()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getEmailSenderName()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableGuestChatter())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableInvitation())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableKnowledgeable())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableNicknameDisplay())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnablePrivateMessages())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableReputation())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableSiteAsContainer())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getEnableTopicAssignmentRules())));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getForgotPasswordTemplate()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getGatherCustomerSentimentData())));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getMaxFileSizeKb()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getNewSenderAddress()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getPicassoSite()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getSelfRegProfile()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getSelfRegistration())));
			excelTemplate.createCell(rowNetwork,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(network.getSendWelcomeEmail())));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getSite()), "UTF-8"));
			
			Enum<NetworkStatus> netStatus = network.getStatus();
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(netStatus.toString()), "UTF-8"));
						
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getUrlPathPrefix()), "UTF-8"));
			excelTemplate.createCell(rowNetwork,cellNum++,URLDecoder.decode(Util.nullFilter(network.getWelcomeTemplate()), "UTF-8"));
				
			//Create Branding Table(ブランド)
			Util.logger.info("Branding started.");
			excelTemplate.createTableHeaders(networkSheet,"NetworkBranding",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			
			//start add by cheng 2017-9-12
			boolean hasbrand = false;
			for(String str:netWorkBrandMap.keySet()){
				if(str.equals(network.getFullName())){
					hasbrand = true;
				}
			}
			//if(network.getBranding() != null){
			if(hasbrand){
				cellNum=1;
				//NetworkBranding brand = network.getBranding();
				NetworkBranding brand = netWorkBrandMap.get(network.getFullName());
			//end add by cheng 2017-9-12
				
				XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getLoginFooterText()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getLoginLogo()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getPageFooter()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getPageHeader()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getPrimaryColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getPrimaryComplementColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getQuaternaryColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getQuaternaryComplementColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getSecondaryColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getTertiaryColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getTertiaryComplementColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getZeronaryColor()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getZeronaryComplementColor()), "UTF-8"));
				//start add by cheng 2017-9-12
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getLoginLogoName()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getLoginRightFrameUrl()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getNetwork()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getStaticLogoImageUrl()), "UTF-8"));
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(brand.getLoginQuaternaryColor()), "UTF-8"));
				//end add by cheng 2017-9-12
			}
			Util.logger.info("Branding completed.");
			
			//Create NavigationLinkSet Table ()
			Util.logger.info("NavigationLinkSet started.");
			excelTemplate.createTableHeaders(networkSheet,"NavigationLinkSet",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getNavigationLinkSet() != null){
				NavigationLinkSet nav = network.getNavigationLinkSet();
				for(NavigationMenuItem navItem : nav.getNavigationMenuItem()){
					cellNum=1;
					XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(navItem.getDefaultListViewId()), "UTF-8"));
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(navItem.getLabel()), "UTF-8"));
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(navItem.getPosition()), "UTF-8"));
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(navItem.getPubliclyAvailable())));
			
					NavigationSubMenu navSub = navItem.getSubMenu();
					if(navSub != null){
						excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(navSub.toString()), "UTF-8"));
					}else{
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(navItem.getSubMenu()));
					}
					
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(navItem.getTarget()), "UTF-8"));
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("NAVIGATION",Util.nullFilter(navItem.getTargetPreference())));
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("NAVIGATIONTYPE",Util.nullFilter(navItem.getType())));
				}			
			}
			Util.logger.info("NavigationLinkSet completed.");
					
			//Create NetworkMemberGroup Table (メンバー)
			Util.logger.info("NetworkMemberGroup started.");
			excelTemplate.createTableHeaders(networkSheet,"NetworkMemberGroup",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getNetworkMemberGroups() != null){
				cellNum=1;
				NetworkMemberGroup net = network.getNetworkMemberGroups();
				XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);
				List<String> strList = new ArrayList<String>();
				for(String str : net.getPermissionSet()){
					strList.add(str);
				}
				if(strList.size() > 0){
					String listString = String.join(", ", strList);
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(listString), "UTF-8"));
				}else{
					excelTemplate.createCell(row,cellNum++,"");
				}
				
				strList.clear();
				for(String str : net.getProfile()){
					strList.add(str);
				}
				if(strList.size() > 0){
					String listString = String.join(", ", strList);
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(listString), "UTF-8"));
				}else{
					excelTemplate.createCell(row,cellNum++,"");
				}
			}
			Util.logger.info("NetworkMemberGroup completed.");
			
			//Create NetworkPageOverride Table ()
			Util.logger.info("NetworkPageOverride started.");
			excelTemplate.createTableHeaders(networkSheet,"NetworkPageOverride",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getNetworkPageOverrides() != null){
				cellNum=1;
				NetworkPageOverride netPage = network.getNetworkPageOverrides();
				XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);
				Enum<NetworkPageOverrideSetting> netSetting = netPage.getChangePasswordPageOverrideSetting();
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(netSetting.toString()), "UTF-8"));
				
				netSetting = netPage.getForgotPasswordPageOverrideSetting();
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(netSetting.toString()), "UTF-8"));
				
				netSetting = netPage.getHomePageOverrideSetting();
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(netSetting.toString()), "UTF-8"));
				
				netSetting = netPage.getLoginPageOverrideSetting();
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(netSetting.toString()), "UTF-8"));
			}
			Util.logger.info("NetworkPageOverride completed.");
			
			//Create ReputationLevelDefinitions Table (評価レベル)
			Util.logger.info("ReputationLevelDefinitions started.");
			excelTemplate.createTableHeaders(networkSheet,"ReputationLevelDefinitions",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getReputationLevels() != null){
				ReputationLevelDefinitions rep = network.getReputationLevels();
				for(ReputationLevel level : rep.getLevel()){
					cellNum=1;
					XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);					
					if(level.getBranding() != null){
						ReputationBranding repBrand = level.getBranding();
						if(repBrand != null){
							excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(repBrand.getSmallImage()), "UTF-8"));
						}
					}else{
						excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(level.getBranding()), "UTF-8"));
					}		
					
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(level.getLabel()), "UTF-8"));
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(level.getLowerThreshold()));
				}
			}
			Util.logger.info("ReputationLevelDefinitions completed.");
			
			//Create ReputationPointsRules Table (評価ポイント)
			Util.logger.info("ReputationPointsRules started.");
			excelTemplate.createTableHeaders(networkSheet,"ReputationPointsRules",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getReputationPointsRules() != null){
				ReputationPointsRules rep = network.getReputationPointsRules();
				for(ReputationPointsRule repPoint : rep.getPointsRule()){
					cellNum=1;
					XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);	
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("EVENTTYPE",Util.nullFilter(repPoint.getEventType())));
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(repPoint.getPoints()));
				}
			}
			Util.logger.info("ReputationPointsRules completed.");
			
			//Create NetworkTabSet Table (タブ)
			Util.logger.info("NetworkTabSet started.");
			excelTemplate.createTableHeaders(networkSheet,"NetworkTabSet",networkSheet.getLastRowNum() + Util.RowIntervalNum);
			if(network.getTabs() != null){
				cellNum=1;
				NetworkTabSet tab = network.getTabs();
				XSSFRow row = networkSheet.createRow(networkSheet.getLastRowNum() + 1);	
				List<String> strList = new ArrayList<String>();
				for(String str : tab.getCustomTab()){
					strList.add(str);
				}
				if(strList.size() > 0){
					String listString = String.join(", ", strList);
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(listString), "UTF-8"));
				}else{
					excelTemplate.createCell(row,cellNum++,"");
				}
				
				excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(tab.getDefaultTab()), "UTF-8"));
				
				strList.clear();
				for(String str : tab.getStandardTab()){
					strList.add(str);
				}
				if(strList.size() > 0){
					String listString = String.join(", ", strList);
					excelTemplate.createCell(row,cellNum++,URLDecoder.decode(Util.nullFilter(listString), "UTF-8"));
				}else{
					excelTemplate.createCell(row,cellNum++,"");
				}				
			}
			Util.logger.info("NetworkTabSet completed.");
			
			//Need to confirm performance issue
			excelTemplate.adjustColumnWidth(networkSheet);
		}
		if(workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalog);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("***no result to export!!!");
		}
		Util.logger.info("ReadNetworkSync End.");
	}
}
