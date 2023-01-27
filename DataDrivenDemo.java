package com.browserLaunch.MavenProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenDemo {
	static WebDriver driver;

	public static void excelData() throws IOException {
		File excel = new File("C:\\Users\\kisho\\eclipse-workspace\\MavenProject\\Data\\website.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		Workbook wb = new XSSFWorkbook(fis);
		Cell cell = wb.getSheetAt(0).getRow(1).getCell(1);
		CellType cellType = cell.getCellType();
		if (cellType.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		} else if (cellType.equals(cellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			System.out.println(numericCellValue);
		}
	}

	public static String excelDataUsingDataFormat() throws IOException {
		File excel = new File("C:\\Users\\kisho\\eclipse-workspace\\MavenProject\\Data\\website.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		Workbook wb = new XSSFWorkbook(fis);
		Row row = wb.getSheetAt(0).getRow(1);
		Cell cellData = wb.getSheetAt(0).getRow(1).getCell(1);
		DataFormatter format = new DataFormatter();
		String formatCellValue = format.formatCellValue(cellData);
		System.out.println(formatCellValue);
		return formatCellValue;
	}

	public static void createSheett() throws IOException, InterruptedException {
		File excel = new File("C:\\Users\\kisho\\eclipse-workspace\\MavenProject\\Data\\website.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet createSheet = wb.createSheet("kishor");
		createSheet.createRow(0).createCell(0).setCellValue("welcome");
		createSheet.getRow(0).createCell(1).setCellValue("to");
		createSheet.createRow(1).createCell(0).setCellValue("java");
		FileOutputStream fos = new FileOutputStream(excel);
		wb.write(fos);
		wb.close();
		System.out.println("===============");
	}

	public static void allData() throws IOException {
		File f = new File("C:\\Users\\kisho\\eclipse-workspace\\MavenProject\\Data\\website.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		int physicalNumberOfRows = wb.getSheetAt(0).getPhysicalNumberOfRows();
		for (int i = 0; i < physicalNumberOfRows; i++) {
			Row row = wb.getSheetAt(0).getRow(i);
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < physicalNumberOfCells; j++) {
				Cell cell = row.getCell(j);
				DataFormatter df = new DataFormatter();
				String formatCellValue = df.formatCellValue(cell);
				System.out.println(formatCellValue);
			}
			System.out.println();
		}
		wb.close();
	}

	public static void dataFromWebsiteToExcel() throws IOException, InterruptedException {
		File f = new File("C:\\Users\\kisho\\eclipse-workspace\\MavenProject\\Data\\website.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.amazon.in/mobile-phones/b/?ie=UTF8&node=1389401031&ref_=nav_cs_mobiles");
		Thread.sleep(5000);
		List<WebElement> findElements = driver
				.findElements(By.xpath("//span[@class='a-size-base-plus a-color-base a-text-normal']"));

		Sheet createSheet = wb.createSheet("mobiles");
		for (int i = 0; i < findElements.size(); i++) {
			WebElement webElement = findElements.get(i);
			String text = webElement.getText();
			Cell createCell = createSheet.createRow(i).createCell(0);
			createCell.setCellValue(text);
		}
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		System.out.println("success");
	}

	public static void main(String[] args) throws IOException, InterruptedException {
		// excelData();
		// excelDataUsingDataFormat();
		// createSheett();
		// allData();
		dataFromWebsiteToExcel();
	}
}
