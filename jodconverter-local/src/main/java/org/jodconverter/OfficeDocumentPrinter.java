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

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import com.sun.star.document.UpdateDocMode;

import org.jodconverter.document.DocumentFormatRegistry;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

public class OfficeDocumentPrinter {

    private final OfficeManager officeManager;
    private final DocumentFormatRegistry formatRegistry;

    private Map<String,?> defaultPrintProperties = createDefaultPrintProperties();

    public OfficeDocumentPrinter(OfficeManager officeManager) {
        this(officeManager, new DefaultDocumentFormatRegistry());
    }

    public OfficeDocumentPrinter(OfficeManager officeManager, DocumentFormatRegistry formatRegistry) {
        this.officeManager = officeManager;
        this.formatRegistry = formatRegistry;
    }

    private Map<String,Object> createDefaultPrintProperties() {
        Map<String,Object> printProperties = new HashMap<String,Object>();
        printProperties.put("Hidden", true);
        printProperties.put("ReadOnly", true);
        printProperties.put("UpdateDocMode", UpdateDocMode.QUIET_UPDATE);
        printProperties.put("Wait", true);
        return printProperties;
    }

    public DocumentFormatRegistry getFormatRegistry() {
        return formatRegistry;
    }

    public void print(File inputFile, String printer, String inputTray, Map<String,?> printProperties) throws OfficeException {
    	Map<String,Object> properties = new HashMap<String,Object>();
    	properties.putAll(defaultPrintProperties);
    	if ( printProperties != null ) {
    		properties.putAll(printProperties);
    	}
    	PrintTask printTask = new PrintTask(inputFile, printer, inputTray, properties);
        printTask.setDefaultLoadProperties(defaultPrintProperties);
    	officeManager.execute(printTask);
    }
}