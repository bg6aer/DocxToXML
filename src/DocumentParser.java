import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;


public class DocumentParser {
	
	private static int rowPos;
	private static int cellPos;
	
	private static List<XWPFTableCell> cells;
	private static List<XWPFTableRow> rows;
	
	private static DocumentWriter xml;
	
	public static void process(File file) throws IOException {
		
		if (file.isFile() && !file.getName().equals(".DS_Store")) {
			
			String path = file.getAbsolutePath().replace("/Users/mike/Downloads/docs/", "/Users/mike/Desktop/output/");			
			
			xml = new DocumentWriter(path);

			xml.writeDocInfo();

			//System.out.println(doc.getAbsolutePath());
			InputStream is = new FileInputStream(file.getAbsolutePath());
			XWPFDocument docx = new XWPFDocument(is);

			List<XWPFTable> tables = docx.getTables();

			XWPFTable table = tables.get(0);

			rows = table.getRows();
			for (rowPos = 0; rowPos < rows.size(); rowPos++) {
				XWPFTableRow row = rows.get(rowPos);

				cells = row.getTableCells();

				for (cellPos = 0; cellPos < cells.size(); cellPos++) {

					String cell = cells.get(cellPos).getText();


					if (!cell.trim().isEmpty() && !cell.equals("Ê")) {
						//System.out.println("-> " + cell);
						parseCell(cell);
					}
				}
			}

			xml.close();
			xml = null;
		} else if (file.isDirectory()) {
			
			boolean createdDir = (new File(file.getAbsolutePath().replace("/Users/mike/Downloads/docs/", "/Users/mike/Desktop/output/")).mkdir());
			
			if (!createdDir) {
				System.out.println("Failed to create a dir " + file.getName());
			}
			
			File[] listFiles = file.listFiles();
			
			if (listFiles != null) {
				for (int i = 0; i < listFiles.length; i++) {
					process(listFiles[i]);
				}
			} else {
				System.out.println("el");
			}
		}
	}
	
	public static void main(String[] args) {
		try {
			File start = new File("/Users/mike/Downloads/docs");
			process(start);
		}
		catch (Exception ex) {
			
		}
	}
	
	private static void nextRow() {
		rowPos++;
		
		cells = rows.get(rowPos).getTableCells();
	}
	
	private static void resetCellPos() {
		cellPos = -1;
	}
	
	private static String requestNextValue() {
		if (cellPos < rows.get(rowPos).getTableCells().size() - 1) {
			return rows.get(rowPos).getCell(++cellPos).getText().replace(":   ", "").replace(":", "").replace("  ", "");
		} else {
			resetCellPos();
			nextRow();
		}
		
		return rows.get(rowPos).getCell(++cellPos).getText().replace(":   ", "").replace(":", "").replace("  ", "");
	}
	
	private static String checkNextValue() {
		if (cellPos < rows.get(rowPos).getTableCells().size() - 1) {
			return rows.get(rowPos).getCell(cellPos + 1).getText().replace(":   ", "").replace("  ", "");
		} else {
			resetCellPos();
			nextRow();
		}
		
		return rows.get(rowPos).getCell(cellPos + 1).getText().replace(":   ", "");
	}
	
	private static String getCurrentValue() {		
		return rows.get(rowPos).getCell(cellPos).getText().replace(":   ", "");
	}
	
	private static void parseCell(String cellText) {
		//System.out.println(rowNum + " " + cellText);
		
		if (cellText.trim().equals("Date :")) {
			// First section!
			xml.writeSection("");
			xml.writeElement(cellText, "");
			xml.writeElement("", requestNextValue());
			xml.writeElement(requestNextValue(), "");
			return;
		}
		else if (cellText.equals("Item Description :")) {
			xml.writeEndSection();
			xml.writeSection(cellText);
			return;
		}
		else if (cellText.equals("Tests :")) {
			xml.writeEndSection();
			xml.writeSection(cellText);
			return;
		}
		else if (cellText.equals("Item") && checkNextValue().equals("Function")) {
			itemFunctionTable();
			return;
		}
		else if (cellText.trim().equals("Test Description")) {
			//testDescTable();
			testDescTable();
			return;
		}
		else if (cellText.trim().equals("Comments:")) {
			xml.writeEndSection();
			xml.writeSection("");
			xml.writeElement("Comments", "");
			return;
		}
		else if (cellText.contains("carried")) {
			xml.writeEndSection();
			xml.writeSection("");
			//writeEndSection();
			xml.writeElement(cellText.replace("          :", ""), "");
			return;
		}
		else if (cellText.trim().contains("Additional Information:")) {
			// Special field
			xml.writeElement(cellText.replace(":", "").replace("ÊÊ", ""), "");
			return;
		}
		else if (cellText.trim().contains("Stamp")) {
			xml.writeElement(cellText.trim().replace("Ê", ""), "stamp", "15", "35", "");
			xml.writeEndSection();
			return;
		}
		else if (cellText.trim().contains("Signature") && checkNextValue().contains("Stamp")) {
			xml.writeElement(cellText.trim(), "");
			return;
		}
		else if (cellText.trim().contains("Applied Value")) {
			appliedTable();
			return;
		}
		xml.writeElement(cellText, requestNextValue());
	}
	
	private static void appliedTable() {
		// 5 cells per row - are usually empty.
		List<XWPFTableRow> tableRows = new ArrayList<XWPFTableRow>();
		
		while(!checkNextValue().contains("carried") && !checkNextValue().contains("Comment")) {
			tableRows.add(rows.get(rowPos));
			nextRow();
		}
		
		for (int i = 0; i < tableRows.size(); i++) {
			// Loop through each row
			XWPFTableRow row = tableRows.get(i);
			List<XWPFTableCell> cells = row.getTableCells();
			
			for (int x = 0; x < cells.size(); x++) {
				if (cells.get(x).getText().contains("Applied Value")) {
					// Skip header fields
					// TODO: imporve by setting headers in dynamic nodes?
					break;
				}
				
				xml.writeRowElement2(cells.get(x).getText(), cells.get(x+1).getText(), cells.get(x+2).getText(), cells.get(x+3).getText(), cells.get(x+4).getText());
				x+=4;
			}
		}
	}
	
	private static void testDescTable() {
		List<XWPFTableRow> tableRows = new ArrayList<XWPFTableRow>();
		
		while(!checkNextValue().contains("carried") && !checkNextValue().contains("Comment")) {
			tableRows.add(rows.get(rowPos));
			nextRow();
		}
		
		for (int i = 0; i < tableRows.size(); i++) {
			// Loop through each row
			XWPFTableRow row = tableRows.get(i);
			List<XWPFTableCell> cells = row.getTableCells();
			
			for (int x = 0; x < cells.size(); x++) {
				// Each cell :-D
				if (cells.get(x).getText().contains("Test Description")) {
					// Skip this row
					break;
				}
				
				xml.writeRowElement(cells.get(x).getText(), cells.get(x+1).getText(), cells.get(x+2).getText(), cells.get(x+3).getText(), cells.get(x+4).getText(), cells.get(x+5).getText(), cells.get(x+6).getText());
				x += 6;
			}
		}
	}
	
	private static void itemFunctionTable() {
//		Item
//		Function
//		Action
//		Date
//		Print/Sign
		List<XWPFTableRow> tableRows = new ArrayList<XWPFTableRow>();
		
		while (!checkNextValue().contains(":") && !checkNextValue().contains("Comment")) {
			
			tableRows.add(rows.get(rowPos));
			nextRow();
		}

		if (cellPos > -1) {
			resetCellPos();
		}
		//System.out.println(checkNextValue());
		
		for (int i = 0; i < tableRows.size(); i++) {
			XWPFTableRow row = tableRows.get(i);
			List<XWPFTableCell> cells = row.getTableCells();
			
			for (int x = 0; x < cells.size(); x++) {
				if (cells.get(x).getText().contains("Item")) {
					// Skip row
					break;
				}
				
				if (i == 1) {
					// There are 6 cells because of that stupid Start Time
					xml.writeRowElement(cells.get(x).getText(), cells.get(x+1).getText(), cells.get(x+3).getText(), cells.get(x + 4).getText(), cells.get(x+5).getText());
					x += 5;
				} else if (i == 2) {
					// Skip this row because it has the fucking Finish Time in
					break;
				}
				else {
					xml.writeRowElement(cells.get(x).getText(), cells.get(x+1).getText(), cells.get(x+2).getText(), cells.get(x+3).getText(), cells.get(x+4).getText());
					x += 4;
				}
			}
		}
	}
	
//	private static void itemFunctionTable_old() {
////		Item
////		Function
////		Action
////		Date
////		Print/Sign
//		nextRow();	// We only want the data :-)
//		resetCellPos();
//		
//		List<String> tableData = new ArrayList<String>();
//		List<String> fields = new ArrayList<String>();
//		
//		while (!checkNextValue().contains("carried") && !checkNextValue().contains("Comment")) {
//			
//			tableData.add(requestNextValue());
//		}
//		
//		for (int i = 0; i < tableData.size(); i++) {
//			if (!tableData.get(i).isEmpty() && !tableData.get(i).contains("Start Time") && !tableData.get(i).contains("Finish Time")) {
//				fields.add(tableData.get(i));
//			}
//		}
//		
//		/* so ugly */
//		for (int i = 0; i < fields.size(); i++) {
//			if (fields.get(i + 1).length() > 1) {
//				writeRowElement(fields.get(i), fields.get(i+1), "", "", "");
//				i += 1;
//			} else {
//				writeRowElement(fields.get(i), "", "", "", "");
//			}
//		}
//	}
}
