package net.karrer.klionewsletter;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Collator;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

/**
 * Read a ExportArtikelFile (csv, "|" as separator, with header) and store:
 * - a list of zwischentitel
 * - the bibliograhic references after auch a title in a TreeSet
 * - reference items are are filled into the template to form a reference line
 * 
 * @author karrer
 *
 */
public class ExportArtikelFile {
	private static String SEPARATOR_REGEX = "\\s*\\|\\s*";
	
	// Buarque, Chico: Mein deutscher Bruder. Roman. S. Fischer Frankfurt a.M. 2016, 256 S., gebunden, 24.- 
	private static String REF_TEMPLATE = "{autor}: {titel}. {verlag}, {zusatz}, {preis2}=={aref}";
	
	List<String> zwischenTitel;
	Map<String, TreeSet<String>> references;
	
	//Path csvFile;
	
	public ExportArtikelFile(Path csvFile) throws IOException {
		//
		zwischenTitel = new ArrayList<String>();
		references = new HashMap<String, TreeSet<String>>();
		Map<String, Integer> itemIndex = new HashMap<String, Integer>();
		List<String> lines = Files.readAllLines(csvFile, StandardCharsets.UTF_8);
		String[] header = null;
		String currentTitel = null;
		for (String line : lines) {
			String[] items = getCSVline(line);
			
			if (items[0].equalsIgnoreCase("aref")) {
				// header line like: "aref"|"autor"|"preis2"|"titel"|"untertitel"|"verlag"|"zusatz"
				// build a map with indices and an array with header names
				header = new String[items.length];
				for (int i = 0; i < items.length; i++) {
					header[i] = items[i];
					itemIndex.put(items[i], new Integer(i));
				}
			} else if (items.length == 1) {
				// Zwischentitel like "Ethnologie"
				currentTitel = items[0];
				zwischenTitel.add(currentTitel);
				references.put(currentTitel, new TreeSet<String>(Collator.getInstance()));
			} else {
				// normal line. Fill the items into the template and store
				String ref = REF_TEMPLATE;
				for (int i = 0; i < items.length; i++) {
					if ("titel".equalsIgnoreCase(header[i])) {
						items[i] = items[i].replaceFirst("\\.\\s*$", "");
					}
					ref = ref.replaceAll("\\{" + header[i] + "\\}", items[i]);
				}
				ref.replaceAll("\\{\\w+\\}", "");
				references.get(currentTitel).add(ref);
			}
		}
	}

	private String[] getCSVline(String line) {
		String[] items = line.split(SEPARATOR_REGEX);
		for (int i = 0; i < items.length; i++) {
			items[i] = items[i].replaceFirst("^\"", "").replaceFirst("\"$", "").replaceAll("\"\"", "\"");
		}
		return items;
	}
	
	/**
	 * Returns the formatted bibliographic references after a Zwischentitel, sorted
	 * 
	 * @param zwischenTitel
	 * @return the sorted references after the zwischentitel
	 */
	public TreeSet<String> getReferences(String zwischenTitel) {
		return references.get(zwischenTitel);
	}

	/**
	 * Returns the list of Zwischentitel in the order as in the input file
	 * @return the list of Zwischentitel 
	 */
	public List<String> getZwischenTitel() {
		return zwischenTitel;
	}
	
	public void printDebugInfo() {
		for (String tit : getZwischenTitel()) {
    		System.err.println(tit+":");
    		TreeSet<String> references = getReferences(tit);
    		for (String ref : references) {
    			System.err.println("    "+ref);
    		}
    	}
	}
	
}

