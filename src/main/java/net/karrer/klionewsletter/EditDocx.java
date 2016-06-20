package net.karrer.klionewsletter;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Locale;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;

import org.xml.sax.SAXException;

public class EditDocx {
	
	private static String TITLE_MARKER = "TITEL";
	private static String REFERENCES_MARKER = "REFS";
	
    public static void main(String [] args) throws IOException, ParserConfigurationException, SAXException, TransformerFactoryConfigurationError, TransformerException {
    	if (args.length < 3) {
    		System.err.println("Usage: $0 in.docx artikelexport.csv out.docx");
    		System.exit(1);
    	}
    	Path inZip = Paths.get(args[0]);
    	Path csvFile = Paths.get(args[1]);
    	Path outZip = Paths.get(args[2]); 
    	
    	// This is used in the Collator for TreeSets
    	Locale.setDefault(new Locale("de", "CH"));
    	
    	// read in the csv file, might have errors
    	ExportArtikelFile artikelFile = new ExportArtikelFile(csvFile);
    	//artikelFile.printDebugInfo();
    	
    	// make a copy of the Vorlage file and work with that
    	Files.copy(inZip, outZip, StandardCopyOption.REPLACE_EXISTING);
    	
    	DocxFile wordFile = new DocxFile(outZip);
    	
    	String docX = wordFile.getDocumentXml();
    	
    	for (String titl : artikelFile.getZwischenTitel()) {
    		docX = replaceTitel(docX, titl);
    		docX = replaceRefs(docX, artikelFile.getReferences(titl), wordFile);
    		System.err.println("Wrote " + titl + ", doc length now=" + docX.length());
    	}
    	
    	wordFile.setDocumentXml(docX);
        
//        String rid = wordFile.addKlioArtikel("123");
//        System.err.println("out.zip has a new elem "+rid);
//        rid = wordFile.addKlioArtikel("124");
//        System.err.println("out.zip has a new elem "+rid);
//        rid = wordFile.addKlioArtikel("125");
//        System.err.println("out.zip has a new elem "+rid);
//        for (int i=1000; i< 1300; i++) {
//        	rid = wordFile.addKlioArtikel(""+i);
//            //System.err.println("out.zip has a new elem "+rid);
//        }
		
        
        wordFile.close();		
    }

	private static String replaceRefs(String docX, TreeSet<String> references, DocxFile wordFile) {
//      <w:p>
//        <w:pPr>
//          <w:pStyle w:val="Normal"/><w:spacing w:before="0" w:after="0"/><w:ind w:left="227" w:right="0" w:hanging="227"/><w:rPr><w:lang w:val="de-CH"/></w:rPr>
//        </w:pPr>
//        <w:hyperlink r:id="rId14">
//          <w:r><w:rPr><w:rStyle w:val="InternetLink"/><w:b/><w:lang w:val="de-CH"/></w:rPr>
//            <w:t>»</w:t>
//          </w:r>
//        </w:hyperlink>
//        <w:r><w:rPr><w:lang w:val="de-CH"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>
//        <w:r><w:rPr><w:lang w:val="de-CH"/></w:rPr>
//          <w:t>REF</w:t>
//        </w:r>
//      </w:p>
		
//      <w:p w:rsidR="00F35C49" w:rsidRPr="00232D54" w:rsidRDefault="004B4CD6" w:rsidP="004C6F9F">
//        <w:pPr><w:ind w:left="227" w:hanging="227"/><w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:rPr><w:lang w:val="de-CH"/></w:rPr>
//        </w:pPr>
//        <w:hyperlink r:id="rId19" w:history="1">
//          <w:r w:rsidR="008E3AEF" w:rsidRPr="00232D54"><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:lang w:val="de-CH"/></w:rPr><w:t>»</w:t></w:r>
//        </w:hyperlink>
//        <w:r w:rsidR="008E3AEF" w:rsidRPr="00232D54"><w:rPr><w:lang w:val="de-CH"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>
//        <w:r w:rsidR="004C6F9F">
//          <w:rPr><w:lang w:val="de-CH"/></w:rPr>
//          <w:t>REFS</w:t>
//        </w:r>
//      </w:p>

		Pattern paraPat = Pattern.compile("<w:p\\b.*?</w:p>");
        Matcher paraMatcher = paraPat.matcher(docX);

	    while (paraMatcher.find()) {
	    	String template = paraMatcher.group();
	    	if (template.contains(">" + REFERENCES_MARKER + "<")) {
	    		String str = docX.substring(0, paraMatcher.start());
	    		String endStr = docX.substring(paraMatcher.end());
	    		System.err.println("  match s:"+paraMatcher.start()+", sz="+template.length()+", e:"+paraMatcher.end());
	    		int l = str.length();
	    		for (String ref : references) {
	    			String[] items = ref.split("=="); // 0 is the reference, 1 is the artikelnummer
	    			String klioArtikelrId = wordFile.addKlioArtikel(items[1]);
	    			str += template.replaceFirst("\"rId\\d+\"", "\"" + klioArtikelrId + "\"")
	    					.replaceFirst(">" + REFERENCES_MARKER + "<", ">" + items[0] + "<");
	    			System.err.println("  Added ref " + klioArtikelrId + " :: " + (str.length()-l) + " chars");
	    			l = str.length();
	    		}
	    		str += endStr;
	    		return str;
	    	}
	    }
	    System.err.println("Cannot find a paragraph with a " + REFERENCES_MARKER + " marker");
	    System.exit(1);
	    return null;
//	    System.err.println("Start:"+matcher.start()+", end:"+matcher.end()
//	    +", g1:"+matcher.group(1).length()+", g2:"+matcher.group(2).length()+", g3:"+matcher.group(3).length());
	}

	private static String replaceTitel(String docX, String titl) {
    	System.err.println("  Added titel " + titl);
		return docX.replaceFirst("<w:t>" + TITLE_MARKER +"</w:t>",  "<w:t>" + titl + "</w:t>");
	}
}
