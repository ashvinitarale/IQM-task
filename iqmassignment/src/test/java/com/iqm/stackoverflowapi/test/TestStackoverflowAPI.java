/**
 * 
 */
package com.iqm.stackoverflowapi.test;

import static io.restassured.RestAssured.given;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.testng.asserts.SoftAssert;

import io.restassured.builder.RequestSpecBuilder;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

/**
 * @author ashvinitarale
 *
 */
public class TestStackoverflowAPI {

	private static RequestSpecification requestSpec;

	@BeforeClass
	public static void createRequestSpecification() {

		requestSpec = new RequestSpecBuilder().setBaseUri("https://api.stackexchange.com").build();
	}

	/**
	 * Test method to verify API Status code for objectSize
	 */
	@Test
	public void requestAndroid_ObjectSize250_checkStatusCode_expectHttp200() {

		given().
			spec(requestSpec).
		when()
			.get("2.2/search/advanced?page=1&pagesize=250&order=desc&sort=activity&title=android&site=stackoverflow\"").
		then()
			.assertThat().statusCode(200);
	}

	/**
	 * Test method to verify content type as JSON
	 */
	@Test
	public void requestAndroid_checkContentType_expectApplicationJson() {

		given().
			spec(requestSpec).
		when()
			.get("/2.2/search/advanced?order=desc&sort=activity&title=android&site=stackoverflow").
		then()
			.assertThat().contentType(ContentType.JSON);
	}


	/**
	 * Method to extract request response
	 * 
	 * @return
	 */
	public Response requestAndroid_fetchfirstObject() {

		return given().
					spec(requestSpec).
			   when().
			   		get("2.2/search/advanced?page=1&order=desc&sort=activity&title=android&site=stackoverflow").
			   then().
			   		extract().response();
	}

	/**
	 * Method to test loading time of link
	 * 
	 * @throws IOException
	 */
	@Test
	public void requestAnroid_measurelodingtime_of_items_link() throws IOException {
		
		// To store link loading time
		long timeInSeconds = 0L;
		
		// To store report dta
		List<List<String>> dataToWrite = new ArrayList<List<String>>();
		
		// call to method which will fetch 250 object for search request android
		Response response = requestAndroid_fetchfirstObject();
		
		// Extracting items link from response
		List<String> itemLinklist = response.jsonPath().getList("items.link");
		
		// Extracting title from response
		List<String> itemTitlelist = response.jsonPath().getList("items.title");
		
		// Iterators for itemLinklist and itemTitlelist
		Iterator<String> itemLinkListItr = itemLinklist.iterator();
		Iterator<String> itemTitlelistItr = itemTitlelist.iterator();
		
		// Using SoftAssert so execution of test method will not stop if any test fail
		SoftAssert softassert = new SoftAssert();
		
		// Computing and asserting loading time
		while (itemLinkListItr.hasNext()) {

			timeInSeconds = given().when().get(itemLinkListItr.next()).timeIn(TimeUnit.SECONDS);

			softassert.assertTrue(timeInSeconds < 4000L);

			// Storing data for report generation
			dataToWrite.add(new ArrayList<String>(Arrays.asList(
							itemTitlelistItr.next() != null ? itemTitlelistItr.next() : " ",
							itemLinkListItr.next() != null ? itemLinkListItr.next() : " ",
							String.valueOf(timeInSeconds), (timeInSeconds < 4000L) ? "Passed" : "Fail")));

		}

		// Method call to generate report in excel file
		writeAPIResultInExcel(System.getProperty("user.dir") + "/src/ExcelReport", "APIExcelReport.xlsx",
				"StackOverFlowAPIResult", dataToWrite);

		// Assert all test
		softassert.assertAll();

	}
	
	/**
	 * Method to write API test result in Excel file
	 * @param filePath
	 * @param fileName
	 * @param sheetName
	 * @param dataToWrite
	 * @throws IOException
	 */
	public void writeAPIResultInExcel(String filePath, String fileName, String sheetName,
			List<List<String>> dataToWrite) throws IOException {
		
		Workbook apiResultWorkbook = null;
		FileInputStream inputStream = null;
		FileOutputStream outputStream = null;

		try {
			
			File newFile = new File(filePath + "/" + fileName);

			// Create an object of FileInputStream class to read excel file
			inputStream = new FileInputStream(filePath + "/" + fileName);

			// Find the file extension by splitting file name in substring and getting only
			// extension name
			String fileExtensionName = fileName.substring(fileName.indexOf("."));

			// Check condition if the file is xlsx file
			if (newFile.exists()) {
				// Load existing
				apiResultWorkbook = WorkbookFactory.create(newFile);
			} else {
				if (fileExtensionName.equals(".xlsx")) {

					// If it is xlsx file then create object of XSSFWorkbook class
					apiResultWorkbook = new XSSFWorkbook(inputStream);

				}
				// Check condition if the file is xls file
				else if (fileExtensionName.equals(".xls")) {

					// If it is xls file then create object of XSSFWorkbook class
					apiResultWorkbook = new HSSFWorkbook(inputStream);

				}
			}

			// Read excel sheet by sheet name
			Sheet sheet = apiResultWorkbook.getSheet(sheetName);

			// Add heading in excel sheet
			Row row = sheet.createRow(0);
			row.createCell(0, CellType.STRING).setCellValue("Title");
			row.createCell(1, CellType.STRING).setCellValue("URL");
			row.createCell(2, CellType.STRING).setCellValue("Load time in seconds");
			row.createCell(3, CellType.STRING).setCellValue("Result");

			// Get cell count
			int cellcount = row.getLastCellNum() - row.getFirstCellNum();

			// Create a loop over the cell of newly created Row
			for (int i = 0; i < dataToWrite.size(); i++) {
				// Add new row
				row = sheet.createRow(i + 1);

				// Fill data in row
				for (int j = 0; j < cellcount; j++) {
					// Add new cell
					Cell cell = row.createCell(j, CellType.STRING);

					cell.setCellValue((dataToWrite.get(i).get(j)));
				}

			}

			// Create an object of FileOutputStream class to create write data in excel file
			outputStream = new FileOutputStream(filePath + "/" + fileName);

			// write data in the excel file
			if (apiResultWorkbook != null)
				apiResultWorkbook.write(outputStream);

		} catch (IOException e) {

			e.printStackTrace();
		} finally {
			
			// Close input stream
			inputStream.close();

			// close output stream
			outputStream.close();

		}

	}

}
