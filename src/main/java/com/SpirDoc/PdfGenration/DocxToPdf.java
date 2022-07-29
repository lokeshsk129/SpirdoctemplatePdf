package com.SpirDoc.PdfGenration;

import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.HeaderFooter;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.TableRow;
import com.spire.doc.ToPdfParameterList;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;

public class DocxToPdf {
	
	public static void main(String[] args) {
		Document document = new Document("D:/EMPLOYEEDETAIL.docx");
		System.out.println(" Load the template document");

		Section section = document.getSections().get(0);
		System.out.println("Get the first section");

		Table table = section.getTables().get(0);
		System.out.println("Get the first table in the section");

		
		
		Map<String, Object> text = new HashMap<String, Object>(); 
		Map<String, String> newMap = new HashMap<String, String>();
		text.put("firstName", "Alex");
		text.put("lastName", "Anderson");
		text.put("mobilePhone", new String[] { "6789765444", "76887544667" });
		text.put("gender", "Male");
		text.put("email", "alexcurg@gmail.com");
		text.put("homeAddress", "4th cross San frnacico USA");
		text.put("dateOfBirth", "10-06-1995");
		for (Map.Entry<String, Object> entry : text.entrySet()) {
			if (entry.getValue() instanceof String) {
				newMap.put(entry.getKey(), (String) entry.getValue());
				
			} 
			

			replaceTextinTable(newMap, table);
			System.out.println("Call the replaceTextinTable method to replace text in table");

			document.saveToFile("D:/MySpirDocx3.docx", FileFormat.Docx_2013);
			System.out.println("Save the result document");

			Document doc2 = new Document();
			System.out.println("Docoment instance creating");

			doc2.loadFromFile("D:/MySpirDocx3.docx");
			System.out.println("docoment is loaded");

			ToPdfParameterList ppl = new ToPdfParameterList();
			System.out.println("pdf instance ir creating");

			ppl.isEmbeddedAllFonts(true);
			System.out.println("embeded all fonts");

			ppl.setDisableLink(true);
			System.out.println("disabled hyperlink");

			doc2.setJPEGQuality(40);
			System.out.println("setting the quality");

			doc2.saveToFile("D:/MySpirDocTopdf.pdf", ppl);
			System.out.println("done");

		}
	}

	// Replace text in table
@SuppressWarnings("unchecked")
	static void replaceTextinTable(Map<String, String> newMap, Table table) {
		for (TableRow row : (Iterable<TableRow>) table.getRows()) {
			for (TableCell cell : (Iterable<TableCell>) row.getCells()) {
				for (Paragraph para : (Iterable<Paragraph>) cell.getParagraphs()) {
					for (Entry<String, String> entry : newMap.entrySet()) {
						
						para.replace("${" + entry.getKey() + "}",entry.getValue(), false, true);
		
					}
				}
			}
		}
	}

}


