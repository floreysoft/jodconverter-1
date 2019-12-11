package org.jodconverter;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.XComponent;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import com.sun.star.view.XPrintable;

import org.jodconverter.office.OfficeUtils;
import org.jodconverter.office.OfficeContext;
import org.jodconverter.office.OfficeException;

public class PrintTask extends AbstractOfficeTask {
	private static Logger logger = Logger.getLogger(PrintTask.class.getName());

	private String printer, inputTray;
	private Map<String, ?> printProperties;

	public PrintTask(File inputFile, String printer, String inputTray,
			Map<String, ?> printProperties) {
		super(inputFile);
		this.printer = printer;
		this.inputTray = inputTray;
		this.printProperties = printProperties;
	}

	@Override
	public void execute(OfficeContext context) throws OfficeException {
		XComponent document = null;
		try {
			document = loadDocument(context, inputFile);
			if (inputTray != null) {
				try {
					XTextDocument aTextDocument = OfficeUtils.cast(XTextDocument.class,
							document);
					XText xText = aTextDocument.getText();
					XTextCursor xTextCursor = xText.createTextCursor();
					XPropertySet xTextCursorProps = OfficeUtils.cast(XPropertySet.class,
							xTextCursor);
					String pageStyleName = xTextCursorProps.getPropertyValue(
							"PageStyleName").toString();
					XStyleFamiliesSupplier xSupplier = OfficeUtils.cast(
							XStyleFamiliesSupplier.class, aTextDocument);
					XNameAccess xFamilies = OfficeUtils.cast(XNameAccess.class,
							xSupplier.getStyleFamilies());
					XNameContainer xFamily = OfficeUtils.cast(XNameContainer.class,
							xFamilies.getByName("PageStyles"));
					XStyle xStyle = OfficeUtils.cast(XStyle.class,
							xFamily.getByName(pageStyleName));
					XPropertySet xStyleProps = OfficeUtils.cast(XPropertySet.class, xStyle);
					XPropertySetInfo propertySetInfo = xStyleProps
							.getPropertySetInfo();
					Property[] properties = propertySetInfo.getProperties();
					boolean trayFound = false;
					for (Property property : properties) {
						if (property.Name.equals("PrinterPaperTray")) {
							trayFound = true;
						}
					}
					if (trayFound) {
						logger.log(Level.INFO,
								"Property PrinterPaperTray found");
						xStyleProps.setPropertyValue("PrinterPaperTray",
								inputTray);
					} else {
						logger.log(Level.SEVERE,
								"Property PrinterPaperTray not found!");
					}
				} catch (Exception officeException) {
					logger.log(Level.SEVERE, "Failed to set input tray="
							+ inputTray + "!", officeException);
				}
			}
			XPrintable printable = OfficeUtils.cast(XPrintable.class, document);
			if (printable != null) {
				Map<String, Object> printerProperty = new HashMap<String, Object>();
				printerProperty.put("Name", printer);
				printable.setPrinter(OfficeUtils.toUnoProperties(printerProperty));
				printable.print(OfficeUtils.toUnoProperties(printProperties));
			}
		} catch (OfficeException officeException) {
			throw officeException;
		} catch (Exception exception) {
			throw new OfficeException("print failed", exception);
		} finally {
			if (document != null) {
				XCloseable closeable = OfficeUtils.cast(XCloseable.class, document);
				if (closeable != null) {
					try {
						closeable.close(true);
					} catch (CloseVetoException closeVetoException) {
						// whoever raised the veto should close the document
					}
				} else {
					document.dispose();
				}
			}
		}
	}
}