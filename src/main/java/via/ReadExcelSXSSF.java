package via;

import java.io.InputStream;
import java.util.Date;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.InputSource;

public class ReadExcelSXSSF {
	/*
	 * So funciona para o ficheiro XLSX
	 * XLSX no fundo é um ZIP com XML's la dentro.
	 * Esta maneira de processar XML é mais rápida e consume menos recurso, mas só dá para ler o XML.
	 * A estrutura é feita para ser lida linha a linha (i.e, le todas as colunas, com valor, da linha 1, etc) 
	 */
	
	public ReadExcelSXSSF() {
		String filename = "C:\\Users\\joao_2\\Desktop\\Book1.xlsx";
		try {
			/* Abrir o ficheiro */
			OPCPackage pkg = OPCPackage.open(filename);
			XSSFReader xssfReader = new XSSFReader(pkg);
			
			/* 
			 * Alguns valores não estão directamente no xml, é preciso ir buscar a sítio comum (as SharedStringTable
			 * Fazem isto para poupar espaco
			 */
			SharedStringsTable sst = xssfReader.getSharedStringsTable();
			
			/*
			 * Abrir o xml correspondente à folha e processa-lo
			 */
			SheetHandler saxHandler = new SheetHandler(sst);
			SAXParserFactory saxFactory = SAXParserFactory.newInstance();
			SAXParser saxParser = saxFactory.newSAXParser();
			
			InputStream sheet1 = xssfReader.getSheet("rId1"); //Deve funcionar com "rId#" or "rSheet#"
			InputSource sheetSource = new InputSource(sheet1);
			saxParser.parse(sheetSource, saxHandler);
			sheet1.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		Date start = new Date();
		new ReadExcelSXSSF();
		Date end = new Date();
		
		long diff = end.getTime() - start.getTime();
		
		long diffSeconds = diff / 1000 % 60;
		long diffMinutes = diff / (60 * 1000) % 60;
		long diffHours = diff / (60 * 60 * 1000) % 24;
		
		System.out.print("Done in: ");
		System.out.print(diffHours + " hours, ");
		System.out.print(diffMinutes + " minutes, ");
		System.out.println(diffSeconds + " seconds.");
	}
	
}
