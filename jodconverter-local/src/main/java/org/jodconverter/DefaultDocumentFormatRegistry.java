//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.jodconverter;

import org.jodconverter.document.DocumentFamily;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.document.SimpleDocumentFormatRegistry;

public class DefaultDocumentFormatRegistry extends SimpleDocumentFormatRegistry {

	public DefaultDocumentFormatRegistry() {
		DocumentFormat pdf = DocumentFormat.builder().name("Portable Document Format").extension("pdf")
				.mediaType("application/pdf").
				storeProperty(DocumentFamily.TEXT, "FilterName", "writer_pdf_Export").
				storeProperty(DocumentFamily.SPREADSHEET, "FilterName", "calc_pdf_Export").
				storeProperty(DocumentFamily.PRESENTATION, "FilterName", "impress_pdf_Export").
				storeProperty(DocumentFamily.DRAWING, "FilterName", "draw_pdf_Export").
				build();
		addFormat(pdf);
//AAA		
		// DocumentFormat swf = DocumentFormat.builder().name("Macromedia Flash").extension("swf")
		// 		.mediaType("application/x-shockwave-flash").build();
		// swf.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "impress_flash_Export"));
		// swf.setStoreProperties(DocumentFamily.DRAWING, Collections.singletonMap("FilterName", "draw_flash_Export"));
		// addFormat(swf);
		
		// // disabled because it's not always available
		// //DocumentFormat xhtml = new DocumentFormat("XHTML", "xhtml", "application/xhtml+xml");
		// //xhtml.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "XHTML Writer File"));
		// //xhtml.setStoreProperties(DocumentFamily.SPREADSHEET, Collections.singletonMap("FilterName", "XHTML Calc File"));
		// //xhtml.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "XHTML Impress File"));
		// //addFormat(xhtml);

		// DocumentFormat html = DocumentFormat.builder().name("HTML").extension("html")
		// 		.mediaType("text/html").build();
        // // HTML is treated as Text when supplied as input, but as an output it is also
        // // available for exporting Spreadsheet and Presentation formats
		// html.setInputFamily(DocumentFamily.TEXT);
		// html.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "HTML (StarWriter)"));
		// html.setStoreProperties(DocumentFamily.SPREADSHEET, Collections.singletonMap("FilterName", "HTML (StarCalc)"));
		// html.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "impress_html_Export"));
		// addFormat(html);
		
		// DocumentFormat odt = DocumentFormat.builder().name("OpenDocument Text").extension("odt").mediaType("application/vnd.oasis.opendocument.text").build();
		// odt.setInputFamily(DocumentFamily.TEXT);
		// odt.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "writer8"));
		// addFormat(odt);

		// DocumentFormat sxw = DocumentFormat.builder().name("OpenOffice.org 1.0 Text Document").extension("sxw")
		// 		.mediaType("application/vnd.sun.xml.writer").build();
		// sxw.setInputFamily(DocumentFamily.TEXT);
		// sxw.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "StarOffice XML (Writer)"));
		// addFormat(sxw);

		// DocumentFormat doc = DocumentFormat.builder().name("Microsoft Word").extension("doc")
		// 		.mediaType("application/msword").build();
		// doc.setInputFamily(DocumentFamily.TEXT);
		// doc.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "MS Word 97"));
		// addFormat(doc);

		// DocumentFormat docx = DocumentFormat.builder().name("Microsoft Word 2007 XML").extension("docx")
		// 		.mediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document").build();
		// docx.setInputFamily(DocumentFamily.TEXT);
        // addFormat(docx);

		// DocumentFormat rtf = DocumentFormat.builder().name("Rich Text Format").extension("rtf")
		// 		.mediaType("text/rtf").build();
		// rtf.setInputFamily(DocumentFamily.TEXT);
		// rtf.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "Rich Text Format"));
		// addFormat(rtf);

		// DocumentFormat wpd = new DocumentFormat("WordPerfect", "wpd", "application/wordperfect");
		// wpd.setInputFamily(DocumentFamily.TEXT);
		// addFormat(wpd);

		// DocumentFormat txt = new DocumentFormat("Plain Text", "txt", "text/plain");
		// txt.setInputFamily(DocumentFamily.TEXT);
		// Map<String,Object> txtLoadAndStoreProperties = new LinkedHashMap<String,Object>();
		// txtLoadAndStoreProperties.put("FilterName", "Text (encoded)");
		// txtLoadAndStoreProperties.put("FilterOptions", "utf8");
		// txt.setLoadProperties(txtLoadAndStoreProperties);
		// txt.setStoreProperties(DocumentFamily.TEXT, txtLoadAndStoreProperties);
		// addFormat(txt);

        // DocumentFormat wikitext = new DocumentFormat("MediaWiki wikitext", "wiki", "text/x-wiki");
        // wikitext.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "MediaWiki"));
        // //addFormat(wikitext);
		
		// DocumentFormat ods = new DocumentFormat("OpenDocument Spreadsheet", "ods", "application/vnd.oasis.opendocument.spreadsheet");
		// ods.setInputFamily(DocumentFamily.SPREADSHEET);
		// ods.setStoreProperties(DocumentFamily.SPREADSHEET, Collections.singletonMap("FilterName", "calc8"));
		// addFormat(ods);

		// DocumentFormat sxc = new DocumentFormat("OpenOffice.org 1.0 Spreadsheet", "sxc", "application/vnd.sun.xml.calc");
		// sxc.setInputFamily(DocumentFamily.SPREADSHEET);
		// sxc.setStoreProperties(DocumentFamily.SPREADSHEET, Collections.singletonMap("FilterName", "StarOffice XML (Calc)"));
		// addFormat(sxc);

		// DocumentFormat xls = new DocumentFormat("Microsoft Excel", "xls", "application/vnd.ms-excel");
		// xls.setInputFamily(DocumentFamily.SPREADSHEET);
		// xls.setStoreProperties(DocumentFamily.SPREADSHEET, Collections.singletonMap("FilterName", "MS Excel 97"));
		// addFormat(xls);

		// DocumentFormat xlsx = new DocumentFormat("Microsoft Excel 2007 XML", "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		// xlsx.setInputFamily(DocumentFamily.SPREADSHEET);
        // addFormat(xlsx);

        // DocumentFormat csv = new DocumentFormat("Comma Separated Values", "csv", "text/csv");
        // csv.setInputFamily(DocumentFamily.SPREADSHEET);
        // Map<String,Object> csvLoadAndStoreProperties = new LinkedHashMap<String,Object>();
        // csvLoadAndStoreProperties.put("FilterName", "Text - txt - csv (StarCalc)");
        // csvLoadAndStoreProperties.put("FilterOptions", "44,34,0");  // Field Separator: ','; Text Delimiter: '"' 
        // csv.setLoadProperties(csvLoadAndStoreProperties);
        // csv.setStoreProperties(DocumentFamily.SPREADSHEET, csvLoadAndStoreProperties);
        // addFormat(csv);

        // DocumentFormat tsv = new DocumentFormat("Tab Separated Values", "tsv", "text/tab-separated-values");
        // tsv.setInputFamily(DocumentFamily.SPREADSHEET);
        // Map<String,Object> tsvLoadAndStoreProperties = new LinkedHashMap<String,Object>();
        // tsvLoadAndStoreProperties.put("FilterName", "Text - txt - csv (StarCalc)");
        // tsvLoadAndStoreProperties.put("FilterOptions", "9,34,0");  // Field Separator: '\t'; Text Delimiter: '"' 
        // tsv.setLoadProperties(tsvLoadAndStoreProperties);
        // tsv.setStoreProperties(DocumentFamily.SPREADSHEET, tsvLoadAndStoreProperties);
        // addFormat(tsv);

		// DocumentFormat odp = new DocumentFormat("OpenDocument Presentation", "odp", "application/vnd.oasis.opendocument.presentation");
		// odp.setInputFamily(DocumentFamily.PRESENTATION);
		// odp.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "impress8"));
		// addFormat(odp);

		// DocumentFormat sxi = new DocumentFormat("OpenOffice.org 1.0 Presentation", "sxi", "application/vnd.sun.xml.impress");
		// sxi.setInputFamily(DocumentFamily.PRESENTATION);
		// sxi.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "StarOffice XML (Impress)"));
		// addFormat(sxi);

		// DocumentFormat ppt = new DocumentFormat("Microsoft PowerPoint", "ppt", "application/vnd.ms-powerpoint");
		// ppt.setInputFamily(DocumentFamily.PRESENTATION);
		// ppt.setStoreProperties(DocumentFamily.PRESENTATION, Collections.singletonMap("FilterName", "MS PowerPoint 97"));
		// addFormat(ppt);

		// DocumentFormat pptx = new DocumentFormat("Microsoft PowerPoint 2007 XML", "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
		// pptx.setInputFamily(DocumentFamily.PRESENTATION);
        // addFormat(pptx);
        
        // DocumentFormat odg = new DocumentFormat("OpenDocument Drawing", "odg", "application/vnd.oasis.opendocument.graphics");
        // odg.setInputFamily(DocumentFamily.DRAWING);
        // odg.setStoreProperties(DocumentFamily.DRAWING, Collections.singletonMap("FilterName", "draw8"));
        // addFormat(odg);
        
        // DocumentFormat svg = new DocumentFormat("Scalable Vector Graphics", "svg", "image/svg+xml");
        // svg.setStoreProperties(DocumentFamily.DRAWING, Collections.singletonMap("FilterName", "draw_svg_Export"));
        // addFormat(svg);
  	}

}
