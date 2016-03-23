/**
 * 
 */
package fr.lynchmaniac.tools.poilight;

import java.io.IOException;

import org.junit.Test;

import fr.lynchmaniac.tools.poilight.entite.Table;

/**
 * @author piard
 * @since 0.1
 *
 */
public class PoiLightTest {


	private static Table data = TestHelper.getTable();
	
	@Test
	public void defaultWorkbook() {
		String excelFilename = "d:\\tmp\\DefaultWorkbook.xlsx";
		PoiLight.generateExcel(excelFilename, data);
		TestHelper.testTable(excelFilename, data);
	}
	
	@Test
	public void defaultStreamingWorkbook() throws IOException {
		String excelFilename = "d:\\tmp\\DefaultStreamingWorkbook.xlsx";
		PoiLight.generateStreamingExcel(excelFilename, data);
		TestHelper.testTable(excelFilename, data);
	}

}
