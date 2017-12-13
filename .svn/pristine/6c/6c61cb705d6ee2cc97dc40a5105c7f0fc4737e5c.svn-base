package wsc;

import java.io.File;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.OutputKeys;  
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

public class WriteXML {
//実行情報をpackage.xmlに保存する
public WriteXML(Map<String,List<String>> map,double ver,String RunTime) throws Exception {
	
   DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
   DocumentBuilder builder = factory.newDocumentBuilder();
   Document doc = builder.newDocument();


   Element root = doc.createElement("Package");
   root.setAttribute("xmlns", "http://soap.sforce.com/2006/04/metadata");
   //version and lastruntime
   Element version = doc.createElement("version");
   Text versionText = doc.createTextNode(String.valueOf(ver));
   version.appendChild(versionText);   
   Element lastRunTime = doc.createElement("lastRunTime");
   Text lastRunTimeText = doc.createTextNode(RunTime);
   lastRunTime.appendChild(lastRunTimeText);
   //types 
   for(String s:map.keySet()){
	   Element types = doc.createElement("types");
	   Element name = doc.createElement("name");
	   Text nameText = doc.createTextNode(s);
	   for(String str:map.get(s)){
		   Element members = doc.createElement("members");      
		   Text membersText = doc.createTextNode(str);  
		   members.appendChild(membersText);
		   types.appendChild(members);
	   }
	   name.appendChild(nameText);
	   types.appendChild(name);
	   root.appendChild(types);
   }
  
   root.appendChild(version);
   root.appendChild(lastRunTime);
   doc.appendChild(root);   
   if(WSC.isBatch){
	   doc2XmlFile(doc,".././conf/package.xml");
   }else{
	   doc2XmlFile(doc,"conf/package.xml");
   }

}

private static boolean doc2XmlFile(Document document, String filename) {
   boolean flag = true;
   try {
	    TransformerFactory tFactory = TransformerFactory.newInstance();
	    Transformer transformer = tFactory.newTransformer();
	    transformer.setOutputProperty(OutputKeys.INDENT, "yes");  
	    // transformer.setOutputProperty(OutputKeys.ENCODING, "GB2312");
	    DOMSource source = new DOMSource(document);
	    StreamResult result = new StreamResult(new File(filename));
	    transformer.transform(source, result);
   } catch (Exception ex) {
	    flag = false;
	    ex.printStackTrace();
   }
   return flag;
}
}