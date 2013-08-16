import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.zip.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class ManualImp {

	//private static String path = "/Users/mike/Desktop/t.docx";
	private static String startDir = "/Users/mike/downloads/docs";
	
	private static String path = "/Users/mike/Desktop/Gauge Calibration and Test Certificate 6m.docx";
	private static StringBuilder sb = new StringBuilder();
	
	public static void main(String args[]) throws ZipException, IOException, SAXException, ParserConfigurationException {
		List<NodeList> rows = new LinkedList<NodeList>();
		
		File tmp = new File(startDir);
		
		FileFilter dirFilter = new FileFilter() {
			public boolean accept(File file) {
				return file.isDirectory();
			}
		};
		
		File[] documentFolders = tmp.listFiles(dirFilter);
		
		for (File documentName : documentFolders) {
			// Get section
			tmp = new File(documentName.getAbsolutePath());
			
			File[] sectionFolders = tmp.listFiles(dirFilter);
			
			for (File sectionName : sectionFolders) {
				// Get items
				tmp = new File(sectionName.getAbsolutePath());
				
				File[] itemFolders = tmp.listFiles(dirFilter);
				
				for (File itemName : itemFolders) {
					tmp = new File(itemName.getAbsolutePath());
					
					File[] docs = tmp.listFiles();
					
					for (File doc : docs) {
						System.out.println(doc.getAbsolutePath());
						if (!doc.getAbsolutePath().contains(".DS_Store")) {
							ZipFile zip = new ZipFile(doc.getAbsolutePath());
							ZipEntry documentXml = zip.getEntry("word/document.xml");
							
							InputStream is = zip.getInputStream(documentXml);
							//System.out.println(is.available());
							
							DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
							Document docx = dbf.newDocumentBuilder().parse(is);
							Element e = docx.getDocumentElement();
							NodeList n = (NodeList)e.getElementsByTagName("w:p");
							
							List<String> texts = new ArrayList<String>();
							
							for (int i = 0; i < n.getLength(); i++) {
								
								Node child = n.item(i);
								String text = child.getTextContent();
								
								if (text != null && !text.trim().isEmpty()  && !text.equals("Ê")) {
										
									// First three elements are always the start section
									texts.add(text);
									System.out.println("NODE[" + i + "] " + text);
								}
							}
							
							// Start parse
							startParse(texts);	
						}
					}
				}
			}
		}
		
		ZipFile zip = new ZipFile(new File(path));
		ZipEntry documentXml = zip.getEntry("word/document.xml");
		
		InputStream is = zip.getInputStream(documentXml);
		//System.out.println(is.available());
		
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		Document doc = dbf.newDocumentBuilder().parse(is);
		Element e = doc.getDocumentElement();
		NodeList n = (NodeList)e.getElementsByTagName("w:p");
		
		List<String> texts = new ArrayList<String>();
		
		for (int i = 0; i < n.getLength(); i++) {
			
			Node child = n.item(i);
			String text = child.getTextContent();
			
			if (text != null && !text.trim().isEmpty()  && !text.equals("Ê")) {
					
				// First three elements are always the start section
				texts.add(text);
				System.out.println("NODE[" + i + "] " + text.trim());
			}
		}
		
		// Start parse
		startParse(texts);
	}
	
	private static void startParse(List<String> data) {
		System.out.println("Parsing document [6], section [60], item [371]");
		System.out.println();
		// we know what the first three is
		dealWithTopSection(data);
		
		int id = 0;
		while (!(data.get(id).trim().equals("Tests :"))) { id++; }
		
		dealWithDescSection(data, id);
		
		id = 0;
		while (!(data.get(id).trim().equals("Comments:")) && !(data.get(id).trim().equals("Test Procedure:"))) { 
			
			if (id < data.size()) {
				id++;
			} else {
				id = data.size() - 1;
				return;
			}
		}
		
		dealWithTestSection(data, id);
		
		System.out.println("<section label=\"\">");
		System.out.println("<field name=\"" + "Comments" + "\" type=\"" + "text" + "\" width=\"150\" height=\"16\" value=\"\" />");
		data.remove(0);
		System.out.println("</section>");
		System.out.println("");
		
		id = 0;//Ê  Company Stamp
		while (!(data.get(id).trim().replaceAll("Ê", "").equals("Company Stamp")) && !(data.get(id).trim().equals("Ê  Company Stamp"))) { id++; }
		dealWithLastSection(data, id);
	}
	
	private static void dealWithGroup(List<String> data, int currentPos, int endPos) {
		
		// Get all of the columns
		// TODO: make it grab headers instead of hard-coded
		currentPos += 5;
		
		//for (int i = currentPos; i < endPos; i++) {
			// First row will have starttime/endtime
			System.out.println("<field type=\"row\" id=\"" + data.get(currentPos) + "\" function=\"" + data.get(currentPos+1) +  "\" starttime=\"\" endtime=\"\" date=\"\" printsign=\"\" />");
		//}
			currentPos += 4;
			
			for (int i = currentPos; i < endPos; i++) {
				String id = data.get(i);
				
				if (data.get(i+1).length() > 1) {
					System.out.println("<field type=\"row\" id=\"" + data.get(i) + "\" function=\"" + data.get(i+1) +  "\" starttime=\"\" endtime=\"\" date=\"\" printsign=\"\" />");
					i += 1;
				}
				else {
					System.out.println("<field type=\"row\" id=\"" + data.get(i) + "\" function=\"" + "" +  "\" starttime=\"\" endtime=\"\" date=\"\" printsign=\"\" />");
				}
			}
	}
	
	private static void dealWithLastSection(List<String> data, int nxtSection) {
		System.out.println("<section label=\"\">");
		
		for (int i = 0; i < nxtSection; i++) {
			
			if (i == 0 || i == 1) {
				// These tables are so inconsistent!
				System.out.println("<field name=\"" + data.get(i).replace("          :", "").replace("ÊÊ", "").replace(":", "") + "\" type=\"" + "text" + "\" width=\"75\" height=\"1\" value=\"\" />");
			} else if (data.get(i).equals("Signature ")) {
				System.out.println("<field name=\"" + data.get(i).replace(" ", "").replace("ÊÊ", "") + "\" type=\"" + "text" + "\" width=\"75\" height=\"1\" value=\"\" />");
			} else {
				System.out.println("<field name=\"" + data.get(i) + "\" value=\"" + data.get(++i).replace(":", "").replace("  ", "") + "\" type=\"text\" width=\"75\" height=\"1\" />");
			}
		}
		
		System.out.println("<field name=\"Company Stamp\" value=\"\" type=\"stamp\" />");
		System.out.println("</section>");
		System.out.println("");
	}
	
	private static void dealWithTable2(List<String> data, int currentPos, int endPos) {
		
		// TODO: not hard-code headers?
		currentPos += 5;
		//System.out.println(data.get(currentPos));
		if (data.get(currentPos).equals(" Comments: ")) {
			System.out.println("<field type=\"row\" id=\"" + "0" + "\" appliedvalue=\"\" rising=\"\" error%=\"\" falling=\"\" error%=\"\" />");
		}
	}
	
	private static void dealWithTable1(List<String> data, int currentPos, int endPos) {
		
	}
	
	private static void dealWithTestSection(List<String> data, int nxtSection) {
		System.out.println("<section label=\"" + data.get(0) + "\">");
		for (int i = 1; i < nxtSection; i++) {
			
			if (data.get(i).equals("Item")) {
				System.out.println("## Data group detected ##");
				// item
				// function
				// action
				// date
				// print/sign
				
				dealWithGroup(data, i, nxtSection);
				break;
			}
			else if (data.get(i).equals("Test Description")) {
//				Test Description
//				Visual Exam & Function Test @ SWL
//				Static Load Test @ 1.25 SWL
//				Static Load Test @ 1.5 SWL
//				Destruct Test
//				NDT
//				Date
				
				dealWithTable1(data, i, nxtSection);
				break;
			}
			else if (data.get(i).equals("Applied Value")) {
//				//Applied Value
//				Rising
//				Error %
//				Falling
//				Error %
				//System.out.println("d");
				dealWithTable2(data, i, nxtSection);
				break;
			}
			else {
				System.out.println("<field name=\"" + data.get(i) + "\" value=\"" + data.get(++i).replace(":", "").replace("  ", "") + "\" type=\"text\" width=\"75\" height=\"1\" />");
			}
		}
		
		for (int i = 0; i < nxtSection; i++) {
			data.remove(0);
		}
		
		System.out.println("</section>");
		System.out.println("");
	}
	
	private static void dealWithDescSection(List<String> data, int nxtSection) {
		
		System.out.println("<section label=\"" + data.get(0) + "\">");
		for (int i = 1; i < nxtSection; i++) {
				System.out.println("<field name=\"" + data.get(i) + "\" label=\"" + data.get(++i).replace(":", "").replace("   ", "") + "\">");
		}
		
		for (int i = 0; i < nxtSection; i++) {
			data.remove(0);
		}
		
		System.out.println("</section>");
		System.out.println("");
	}
	
	private static void dealWithTopSection(List<String> data) {
		System.out.println("<section label=\"\">");
		System.out.println("<field name=\"" + data.get(0).replace(" :", "") + "\" type=\"text\" width=\"24\" height=\"1\" value=\"\" />");
		
		if (!data.get(2).trim().equals("N¡")) {
			System.out.println("<field name=\"\" type=\"text\" width=\"24\" height=\"1\" value=\"" + data.get(1) + data.get(2) + "\" />");
			System.out.println("<field name=\"" + data.get(3) + "\" type=\"text\" width=\"24\" height=\"1\" value=\"" + "" + "\" />");
			data.remove(0);
			data.remove(0);
			data.remove(0);
			data.remove(0);
		} else {
		
			System.out.println("<field name=\"\" type=\"text\" width=\"24\" height=\"1\" value=\"" + data.get(1) + "\" />");
			System.out.println("<field name=\"" + data.get(2) + "\" type=\"text\" width=\"24\" height=\"1\" value=\"" + "" + "\" />");
			data.remove(0);
			data.remove(0);
			data.remove(0);
		}
	}
	
}
