package net.karrer.klionewsletter;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * Edit a *.docx (Word) file in-place.
 * - add hyperlinks to the relationships file
 * - read and write the document.xml file
 * 
 * @author karrer
 *
 */
public class DocxFile {
	FileSystem zipfs;
	Path wordFile;
	Path docXmlPath;
	String documentXml;
	Path relsFilePath;
	Document relsDoc;
	Set<String> usedIds;

	/** Constructor
	 * 
	 * @param wordFile
	 * @throws ParserConfigurationException
	 * @throws IOException
	 * @throws SAXException
	 */
	public DocxFile(Path wordFile) throws ParserConfigurationException, IOException, SAXException {
		this.wordFile = wordFile;
		
		  // Open the word file as a zip filesystem	
//        Map<String, String> env = new HashMap<>(); 
//        // env.put("create", "true");
//        URI uri = URI.create("jar:file:" + wordFile);
//        zipfs = FileSystems.newFileSystem(uri, env);
        
        //Path zipfile = Paths.get("/codeSamples/zipfs/zipfstest.zip");
        zipfs = FileSystems.newFileSystem(wordFile, null);
        
        // read the document.xml file into a string
        docXmlPath = zipfs.getPath("word/document.xml");
        documentXml = new String(Files.readAllBytes(docXmlPath), StandardCharsets.UTF_8);
        System.err.println("Read document.xml, " + documentXml.length() + " chars");

		// read and parse the Relationships file
        relsFilePath = zipfs.getPath("/word/_rels/document.xml.rels");
        DocumentBuilder docBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
		InputStream relsFile = Files.newInputStream(relsFilePath);
		relsDoc = docBuilder.parse(relsFile);
		relsFile.close();
		
		// Get the root element <Relationships> and its children
		//   <Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
		//                 Target="http://www.klio-buch.ch/artikel_647.ahtml" TargetMode="External"/>

		Node relationships = relsDoc.getFirstChild();
		NodeList relations = relationships.getChildNodes();
		usedIds = new HashSet<String>();
		for (int i = 0; i < relations.getLength(); i++) {
			Node node = relations.item(i);
			if (node != null) {
	            Element rel = (Element) node;
	            String id = rel.getAttribute("Id");
	            usedIds.add(id);
	            //System.out.println("used id:"+id);
			}
		}
		System.err.println("Rels file has used " + relations.getLength() + " rIds");
	}
	
	public String getDocumentXml() {
		return documentXml;
	}

	public void setDocumentXml(String documentXml) {
		this.documentXml = documentXml;
	}

	/**
	 * Inserts a hyperlink reference to a www.klio-buch.ch artikel into the relationships file
	 * and returns the rid string (of the form "rId1234").
	 * 
	 * @param artikelNr
	 * @return a new rId that is not yet present in the relationships file
	 */
	public String addKlioArtikel(String artikelNr) {
		Node relationships = relsDoc.getFirstChild();
		int num = 1;
		while (usedIds.contains("rId"+num)) {
			num++;
		}
		String rid = "rId"+num;
		Element newRel = relsDoc.createElement("Relationship");
		newRel.setAttribute("Id", rid);
		newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
		newRel.setAttribute("Target", "http://www.klio-buch.ch/artikel_" + artikelNr + ".ahtml");
		newRel.setAttribute("TargetMode", "External");
		relationships.appendChild(newRel);
		usedIds.add(rid);
		return rid;
	}
	
	/**
	 * Debug routine that returns the possible edited relationships file in a formatted xml string.
	 * @return the xml as a string
	 * @throws TransformerFactoryConfigurationError
	 * @throws TransformerException
	 */
	public String getRelationshipXml() throws TransformerFactoryConfigurationError, TransformerException {
		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		//Result output = new StreamResult(os);
		Source input = new DOMSource(relsDoc);
		
		StringWriter sw = new StringWriter();
	    StreamResult sr = new StreamResult(sw);
	    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
	    transformer.transform(input, sr);
	    return sw.toString();
	}
	
	/**
	 * Writes the (possibly edited) document.xml and refs file to the zip archive and closes it.
	 * 
	 * @throws IOException
	 * @throws TransformerFactoryConfigurationError
	 * @throws TransformerException
	 */
	public void close() throws IOException, TransformerFactoryConfigurationError, TransformerException  {
		
		// first delete the old, then write out the neww doxument.xml file
		Files.delete(docXmlPath);
		
    BufferedWriter writer = Files.newBufferedWriter(docXmlPath, StandardCharsets.UTF_8);
    writer.write(documentXml);
    writer.close();
		
		
		// first delete and then write out the relationships file, using the identity transformer
		// indent for readability
		Files.delete(relsFilePath);
		
		OutputStream outputStream = Files.newOutputStream(relsFilePath);
		
		// Use a Transformer for output
		Source dom = new DOMSource(relsDoc);
    Transformer transformer = TransformerFactory.newInstance().newTransformer();
    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
    StreamResult result = new StreamResult(outputStream);
    transformer.transform(dom, result);
		outputStream.close();
		
		// close the zipfs -- this writes out the *.docx file
		zipfs.close();
	}

}
