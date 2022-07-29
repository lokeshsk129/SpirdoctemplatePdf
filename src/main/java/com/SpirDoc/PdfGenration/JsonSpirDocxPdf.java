package com.SpirDoc.PdfGenration;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
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

public class JsonSpirDocxPdf {

	static String imgFile1 = "D:/card.jpg";
	static String resourcePath1 = "D:/EMPLOYEEDETAIL1.docx";
	static String docPath1 = "D:/MySpirDocx3.docx";
	static String pdfPath1 = "D:/MySpirDocTopdf.pdf";
	static String jsonFile = "D:/testJsonFile.json";

	public static void main(String[] args) throws JsonParseException, JsonMappingException, IOException, NullPointerException {
 try {
		Document document = new Document(resourcePath1);
		System.out.println(" Load the template document");

		Section section = document.getSections().get(0);
		System.out.println("Get the first section");

		Table table = section.getTables().get(0);
		System.out.println("Get the first table in the section");

		ObjectMapper mapper = new ObjectMapper();

		File jsonfile = new File(jsonFile);

		Map<String, Object> jsonMap = new HashMap<String, Object>();
		Map<String, String> newJsonMap = new HashMap<String, String>();
		for (Map.Entry<String, Object> entry : jsonMap.entrySet()) {
			if (entry.getValue() instanceof String) {
				newJsonMap.put(entry.getKey(), (String) entry.getValue());
			}
		}
		newJsonMap = mapper.readValue(jsonfile, new TypeReference<Map<String, String>>() {
		});

		replaceTextinTable(newJsonMap, table);
		System.out.println("Call the replaceTextinTable method to replace text in table");

		replaceTextWithImage(document, "avatar", "D:/card.jpg");
		System.out.println("Call the replaceTextWithImage method to replace text with image");

		document.saveToFile(docPath1, FileFormat.Docx_2013);
		System.out.println("Save the result document");

		Document doc2 = new Document();
		System.out.println("Docoment instance creating");

		doc2.loadFromFile(docPath1);
		System.out.println("docoment is loaded");

		ToPdfParameterList ppl = new ToPdfParameterList();
		System.out.println("pdf instance ir creating");

		ppl.isEmbeddedAllFonts(true);
		System.out.println("embeded all fonts");

		ppl.setDisableLink(true);
		System.out.println("disabled hyperlink");

		doc2.setJPEGQuality(40);
		System.out.println("setting the quality");

		doc2.saveToFile(pdfPath1, ppl);
		System.out.println("filed saved successfully");
       }
   catch(Exception e) {
	   e.printStackTrace();
	   
   }

	}

	// Replace text in table
	@SuppressWarnings(value = { "unchecked" })
	static void replaceTextinTable(Map<String, String> newJsonMap, Table table) {
		for (TableRow row : (Iterable<TableRow>) table.getRows()) {
			for (TableCell cell : (Iterable<TableCell>) row.getCells()) {
				for (Paragraph para : (Iterable<Paragraph>) cell.getParagraphs()) {
					for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
						para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);

					}
				}
			}
		}
	}

	// Replace text with image
	static void replaceTextWithImage(Document document, String stringToReplace, String imagePath) throws NullPointerException {
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
	static void replaceTextinDocumentBody(Map<String, String> newJsonMap, Document document) {
		for (Section section : (Iterable<Section>) document.getSections()) {
			for (Paragraph para : (Iterable<Paragraph>) section.getParagraphs()) {
				for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
					para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
				}
			}
		}
	}

	// Replace text in header or footer
	@SuppressWarnings("unchecked")
	static void replaceTextinHeaderorFooter(Map<String, String> newJsonMap, HeaderFooter headerFooter) {
		for (Paragraph para : (Iterable<Paragraph>) headerFooter.getParagraphs()) {
			for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
				para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
			}
		}
	}

}
