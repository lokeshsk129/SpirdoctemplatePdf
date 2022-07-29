package com.SpirDoc.PdfGenration;

import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;



public class SpricDocPdf {

	public static void main(String[] args) throws IOException, NoClassDefFoundError {
        
		Document document = new Document("D:/EMPLOYEEDETAIL1.docx");
		System.out.println(" Load the template document");

		Section section = document.getSections().get(0);
		System.out.println("Get the first section");

		Table table = section.getTables().get(0);
		System.out.println("Get the first table in the section");

		// Create a map of values for the template
		Map<String, Map<String, String>> data = new HashMap<String, Map<String, String>>();
	    Map<String, String> text = new HashMap<String, String>();
	    text.put("firstName", "Alex");
	    text.put("lastName", "Anderson");
	    text.put("gender", "Male");
	    text.put("mobileNum", "9972464834");
	    text.put("officeNum", "9972464834");
	    text.put("homeNum", "0848267283");
	    text.put("email","alexAnderson@gmail.com");
	    text.put("homeAddress", "123 High Street");
	    text.put("dateOfBirth", "6th June, 1986");
	    text.put("education", "University of South Florida, September 2013 - June 2017");
	    text.put("employmentHistory", "Automation Inc. November 2013 - Present");   
	    data.put("text", text);
		
	    
	    
		
		replaceTextinTable(text, table);     
		System.out.println("Call the replaceTextinTable method to replace text in table");

		replaceTextWithImage(document, "avatar", "D:/card.jpg");
		System.out.println("Call the replaceTextWithImage method to replace text with image");

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
		System.out.println("filed saved successfully");

	}

	// Replace text in table
	@SuppressWarnings(value = { "unchecked" })
	static void replaceTextinTable(Map<String, String> data, Table table) {
		for (TableRow row : (Iterable<TableRow>) table.getRows()) {
			for (TableCell cell : (Iterable<TableCell>) row.getCells()) {
				for (Paragraph para : (Iterable<Paragraph>) cell.getParagraphs()) {
					for (Map.Entry<String, String> entry : data.entrySet()) {
						para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
					
						
					}
				}
			}
		}
	}
	// Replace text with image
	static void replaceTextWithImage(Document document, String stringToReplace, String imagePath) {
		TextSelection[] selections = document.findAllString("${" + stringToReplace + "}", false, true);
		int index = 0;
		TextRange range = null;
		for (Object obj : selections) { 
			TextSelection textSelection = (TextSelection) obj; // Creates a text selection for the given range.
			DocPicture pic = new DocPicture(document); // Initializes a new instance of the DocPicture class.
			pic.loadImage(imagePath); // image is loading
			pic.setWidth(160);
			pic.setHeight(120);
			range = textSelection.getAsOneRange();
			index = range.getOwnerParagraph().getChildObjects().indexOf(range);
			range.getOwnerParagraph().getChildObjects().insert(index, pic);
			range.getOwnerParagraph().getChildObjects().remove(range);
		}
	}
	

	@SuppressWarnings("unchecked")
	static void replaceTextinDocumentBody(Map<String, String> data, Document document) {
		for (Section section : (Iterable<Section>) document.getSections()) {
			for (Paragraph para : (Iterable<Paragraph>) section.getParagraphs()) {
				for (Map.Entry<String, String> entry : data.entrySet()) {
					para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
				}
			}
		}
	}

	// Replace text in header or footer
	@SuppressWarnings("unchecked")
	static void replaceTextinHeaderorFooter(Map<String, String> data, HeaderFooter headerFooter) {
		for (Paragraph para : (Iterable<Paragraph>) headerFooter.getParagraphs()) {
			for (Map.Entry<String, String> entry : data.entrySet()) {
				para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
			}
		}
	}

}
