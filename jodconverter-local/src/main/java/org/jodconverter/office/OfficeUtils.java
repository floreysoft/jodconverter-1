/*
 * Copyright 2004 - 2012 Mirko Nasato and contributors
 *           2016 - 2019 Simon Braconnier and contributors
 *
 * This file is part of JODConverter - Java OpenDocument Converter.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.jodconverter.office;

import java.io.File;
import java.util.Map;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.UnoRuntime;

import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

/** Provides helper functions for office. */
public final class OfficeUtils {

  /**
   * Stops an <code>OfficeManager</code> unconditionally.
   *
   * <p>Equivalent to {@link OfficeManager#stop()}, except any exceptions will be ignored. This is
   * typically used in finally blocks.
   *
   * <p>Example code:
   *
   * <pre>
   * OfficeManager manager = null;
   * try {
   *     manager = LocalOfficeManager().make();
   *     manager.start();
   *
   *     // process manager
   *
   * } catch (Exception e) {
   *     // error handling
   * } finally {
   *     OfficeUtils.stopQuietly(manager);
   * }
   * </pre>
   *
   * @param manager the manager to stop, may be null or already stopped.
   */
  public static void stopQuietly(final OfficeManager manager) {

    try {
      if (manager != null) {
        manager.stop();
      }
    } catch (final OfficeException ex) { // NOSONAR
      // ignore
    }
  }

  // Suppresses default constructor, ensuring non-instantiability.
  private OfficeUtils() {
    throw new AssertionError("Utility class must not be instantiated");
  }

  public static <T> T cast(Class<T> type, Object object) {
    return (T) UnoRuntime.queryInterface(type, object);
  }

  public static PropertyValue property(String name, Object value) {
    PropertyValue propertyValue = new PropertyValue();
    propertyValue.Name = name;
    propertyValue.Value = value;
    return propertyValue;
  }

  @SuppressWarnings("unchecked")
  public static PropertyValue[] toUnoProperties(Map<String, ?> properties) {
    PropertyValue[] propertyValues = new PropertyValue[properties.size()];
    int i = 0;
    for (Map.Entry<String, ?> entry : properties.entrySet()) {
      Object value = entry.getValue();
      if (value instanceof Map) {
        Map<String, Object> subProperties = (Map<String, Object>) value;
        value = toUnoProperties(subProperties);
      }
      propertyValues[i++] = property((String) entry.getKey(), value);
    }
    return propertyValues;
  }

  public static String toUrl(File file) {
    String path = file.toURI().getRawPath();
    String url = path.startsWith("//") ? "file:" + path : "file://" + path;
    return url.endsWith("/") ? url.substring(0, url.length() - 1) : url;
  }
}
