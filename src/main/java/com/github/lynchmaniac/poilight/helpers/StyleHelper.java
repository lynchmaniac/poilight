package com.github.lynchmaniac.poilight.helpers;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * This class contains all the technical utilities related to styles.
 * 
 * @author vpiard
 * @since 0.1
 */
public final class StyleHelper {

  
  private StyleHelper() {
    throw new IllegalStateException("Utility class");
  }
  
  /**
   * Returns the final color from the three red, green and blue.
   * 
   * @param red the hexadecimal value of the red
   * @param green the hexadecimal value of the green 
   * @param blue the hexadecimal value of the blue
   * @return the final color
   */
  public static XSSFColor getColor(int red, int green, int blue) {
    byte[] rgb = new byte[3];
    rgb[0] = (byte) red; 
    rgb[1] = (byte) green; 
    rgb[2] = (byte) blue; 

    return new XSSFColor(rgb, new DefaultIndexedColorMap());

  }

}
