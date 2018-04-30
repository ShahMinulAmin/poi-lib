package sajib.test.poi;

import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class XSSFReaderExample {
	public static void main(String[] args) throws Exception {
		final String FILE_NAME = "./xssf_example.xlsx";
		XSSFReaderExample example = new XSSFReaderExample();
		example.readExcelFile(FILE_NAME);
	}

	public void readExcelFile(String filename) throws Exception {
		OPCPackage opcPackage = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(opcPackage);
		SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
		XMLReader parser = getSheetParser(sharedStringsTable);

		Iterator<InputStream> sheets = xssfReader.getSheetsData();
		while (sheets.hasNext()) {
			System.out.println("Processing sheet:");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			System.out.println();
		}
	}

	public XMLReader getSheetParser(SharedStringsTable sharedStringsTable) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		ContentHandler handler = new SheetHandler(sharedStringsTable);
		parser.setContentHandler(handler);
		return parser;
	}

	/** sheet handler class for SAX2 events */
	private static class SheetHandler extends DefaultHandler {
		private SharedStringsTable sharedStringsTable;
		private String contents;
		private boolean isCellValue;
		private boolean fromSST;

		private SheetHandler(SharedStringsTable sharedStringsTable) {
			this.sharedStringsTable = sharedStringsTable;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// Clear contents cache
			contents = "";				
			// element row represents Row
			if (name.equals("row")) {
				String rowNumStr = attributes.getValue("r");
				System.out.println("Row# " + rowNumStr);
			}		
			// element c represents Cell
			else if (name.equals("c")) {
				// attribute r represents the cell reference
				System.out.print(attributes.getValue("r") + " - ");
				// attribute t represents the cell type
				String cellType = attributes.getValue("t");
				if (cellType != null && cellType.equals("s")) {
					// cell type s means value will be extracted from SharedStringsTable
					fromSST = true;
				} 
			// element v represents value of Cell	
			} else if (name.equals("v")) {
				isCellValue = true;
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			if (isCellValue) {
				contents += new String(ch, start, length);	
			}
		}		
		
		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			if (isCellValue && fromSST) {
				int index = Integer.parseInt(contents);
				contents = new XSSFRichTextString(sharedStringsTable.getEntryAt(index)).toString();
				System.out.println(contents);
				isCellValue = false;
				fromSST = false;
			}
		}
	}
}
