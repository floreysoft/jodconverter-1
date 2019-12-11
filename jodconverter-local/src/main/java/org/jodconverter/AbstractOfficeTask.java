package org.jodconverter;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.task.ErrorCodeIOException;

import org.jodconverter.document.DocumentFormat;
import org.jodconverter.office.OfficeContext;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;
import org.jodconverter.task.OfficeTask;

public abstract class AbstractOfficeTask implements OfficeTask {
    protected final File inputFile;
    private Map<String,?> defaultLoadProperties;
    private DocumentFormat inputFormat;

    public AbstractOfficeTask(File inputFile) {
		this.inputFile = inputFile;
	}

	public void setInputFormat(DocumentFormat inputFormat) {
        this.inputFormat = inputFormat;
    }
    
    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    protected Map<String,?> getLoadProperties(File inputFile) {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    protected XComponent loadDocument(OfficeContext context, File inputFile) throws OfficeException {
        if (!inputFile.exists()) {
            throw new OfficeException("input document not found");
        }
        XComponentLoader loader = OfficeUtils.cast(XComponentLoader.class, context.getDesktopService());
        Map<String,?> loadProperties = getLoadProperties(inputFile);
        XComponent document = null;
        try {
            document = loader.loadComponentFromURL(OfficeUtils.toUrl(inputFile), "_blank", 0, OfficeUtils.toUnoProperties(loadProperties));
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load document: " + inputFile.getName(), illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load document: "  + inputFile.getName() + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not load document: "  + inputFile.getName(), ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load document: "  + inputFile.getName());
        }
        return document;
    }
}