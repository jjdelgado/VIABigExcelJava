package via;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/*
 * Basicamente esta forma de processar XML consiste em ler o XML linha a linha e enviar eventos para o utilizador.
 * Os eventos que apanho são:
 * * startElement - sempre que se abre uma tag
 * * endElement - sempre que se fecha uma tag
 * * characters - lê o que esta entre tags
 */
public class SheetHandler extends DefaultHandler {
	private SharedStringsTable sst;
	private String string;
	private boolean nextValueIsSharedString;
	
	public SheetHandler(SharedStringsTable sst) {
		this.sst = sst;
	}
	
	@Override
	public void startElement(String uri, String localName, String qName,
			Attributes attributes) throws SAXException {
		/*
		 * Aqui vejo se se abre a tag <c>, que corresponde a uma celula em excel
		 */
		if (qName.equals("c")) {
			String cellType =  attributes.getValue("t"); // obtenho o tipo da celula (atributo "t")
			System.out.print(attributes.getValue("r") + "(type: " + cellType + ") - "); // O atributo "r" é a coordenada da celula - A1, A2, B7, etc
			if (cellType != null && cellType.equals("s")) { // se o tipo da celula for "s" - significa que tens de ler da tal SharedStringTable - no fim do ficheiro estao todos os tipo
				nextValueIsSharedString = true;
			}
			else {
				nextValueIsSharedString = false;
			}
		}
		
		string = ""; 
	}
	
	@Override
	public void endElement(String uri, String localName, String qName)
			throws SAXException {
		
		/*
		 * basicamente aqui como sei que é para ler da tabela, obtenho o index e vou buscar o valor
		 */
		if (nextValueIsSharedString) {
			int idx = Integer.parseInt(string);
			string = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
		}
		
		/*
		 * a tag <v> significa valor de uma celula
		 */
		if (qName.equals("v")) {
			System.out.println(string);
		}
	}
	
	@Override
	public void characters(char[] ch, int start, int length)
			throws SAXException {
		string += new String(ch, start, length);
	}
}

/*
 * 
Valid value	Description
b			Boolean
n			Number
e			Error
s			Shared String
str			String
inlineStr	Inline String
*/