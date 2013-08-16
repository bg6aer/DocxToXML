import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;


public class DocumentWriter {
	
	private PrintWriter xml;
	private String path;
	
	public DocumentWriter(String path) {
		this.path = path;
		
		try {

			path = path.replace(".docx", "");
			xml = new PrintWriter(new FileWriter(new File(path)));
			
			xml.println("<xml version=\"1.0\">");
			xml.println("	<document>");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void close() {
		try {
			xml.println("		<section label=\"\">");
			xml.println("			<field name=\"\" type=\"label\" value=\"IMCA Category Competent Person:\n1 Ð Diving or Life Support Supervisor\n2 Ð Diving Technician or Chief Engineer\n3 Ð Classification Society or in-house chartered Engineer\n4 Ð Manufacturer or Supplier of the Equipment\" />");
			xml.println("		</section>");
			xml.println("");
			
			xml.println("	</document>");
			xml.println("</xml>");
			
			//xml.flush();
			xml.close();
		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}
	}
	
	public void writeDocInfo() {
		
		String[] split = path.split("/");
		
		xml.println("		<docinfo>");
		xml.println("			<document>" + split[5] + "</document>");
		xml.println("			<section>" + split[6] + "</section>");
		xml.println("			<item>" + split[7] + "</item>");
		xml.println("			<certificate>" + split[8] + "</certificate>");
		xml.println("		</docinfo>");
		xml.println("");
	}
	
	public void writeEndSection() {
		xml.println("		</section>");
		xml.println();
	}
	public void writeSection(String name) {
		xml.println("		<section label=\"" + name + "\">");
	}
	
	public void writeRowElement2(String appliedval, String rising, String error1, String falling, String error2) {
		xml.println("			<field type=\"row\" appliedvalue=\"" + appliedval + "\" rising=\"" + rising + "\" errorPercent1=\"" + error1 + "\" falling=\"" + falling + "\" errorPercent2=\"\" />");
	}
	
	public void writeElement(String name, String type, String width, String height, String value) {
		xml.println("			<field name=\"" + name + "\" type=\"" + type + "\" width=\"" + width + "\" height=\"" + height + "\" value=\"" + value + "\" />");
	}
	public void writeRowElement(String testdesc, String functiontest, String statictest, String statictest2, String destruct, String ndt, String date) {
		xml.println("			<field type=\"row\" testdesc=\"" + testdesc + "\" statictest=\"" + statictest + "\" statictest2=\"" + statictest2 + "\" destruct=\"" + destruct + "\" ndt=\"" + ndt + "\" date=\"" + date + "\" />");
	}
	public void writeRowElement(String id, String function, String action, String date, String printSign) {
		xml.println("			<field type=\"row\" id=\"" + id + "\" function=\"" + function + "\" action=\"" + action + "\" starttime=\"\" endtime=\"\" date=\"" + date + "\" printsign=\"\" />");
	}
	
	public void writeElement(String name, String value) {
		xml.println("			<field name=\"" + name + "\" type=\"text\" width=\"75\" height=\"1\" value=\"" + value + "\" />");
	}
}
