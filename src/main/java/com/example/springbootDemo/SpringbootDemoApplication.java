package com.example.springbootDemo;

import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@SpringBootApplication
@RestController
public class SpringbootDemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(SpringbootDemoApplication.class, args);
	}
	
	@RequestMapping("/")
	public String home() {
		return "Hello Spring Boot!";
	}

	@GetMapping("/hello")
	public String hello(HttpServletResponse response, @RequestParam(value = "name", defaultValue = "World") String name) throws IOException {
		return String.format("Hello %s!", name);
	}
	
	@GetMapping("/getExcel")
	public String excel(HttpServletResponse response, @RequestParam(value = "name", defaultValue = "World") String name) throws IOException {
		System.out.println("hello from hello method");
		
		Workbook wb = new SXSSFWorkbook(100);
		Sheet nrRequestSheet = wb.createSheet("Non Request");
		nrRequestSheet.addMergedRegion(new CellRangeAddress(0,0,0,8));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(1,1,1,3));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(1,1,7,8));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(2,2,1,3));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(2,2,4,5));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(2,2,7,8));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(3,3,1,3));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(3,3,4,5));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(4,4,1,5));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(5,5,1,5));
		nrRequestSheet.addMergedRegion(new CellRangeAddress(5,5,7,8));
		
		nrRequestSheet.createRow(0).createCell(0).setCellValue("Non Routine Test Request");
		
		Row row1 = nrRequestSheet.createRow(1);
		row1.createCell(0).setCellValue("Batch Number");
		row1.createCell(1).setCellValue("");
		row1.createCell(4).setCellValue("Sample Date");
		row1.createCell(5).setCellValue("");
		row1.createCell(6).setCellValue("Workarea ID");
		row1.createCell(7).setCellValue("");
		
		Row row2 = nrRequestSheet.createRow(2);
		row2.createCell(0).setCellValue("Client Code");
		row2.createCell(1).setCellValue("");
		row2.createCell(2).setCellValue("");
		row2.createCell(6).setCellValue("Alternate Contact");
		row2.createCell(7).setCellValue("");
		row2.createCell(8).setCellValue("");
		
		Row row3 = nrRequestSheet.createRow(3);
		row3.createCell(0).setCellValue("Requestor");
		row3.createCell(1).setCellValue("");
		row3.createCell(2).setCellValue("");
		row3.createCell(6).setCellValue("Non Std Testing");
		row3.createCell(7).setCellValue("");
		row3.createCell(8).setCellValue("NST Number");
		row3.createCell(9).setCellValue("");
		
		Row row4 = nrRequestSheet.createRow(4);
		row4.createCell(0).setCellValue("Business Unit");
		row4.createCell(1).setCellValue("");
		row4.createCell(6).setCellValue("Project");
		row4.createCell(7).setCellValue("");
		row4.createCell(8).setCellValue("Project Number");
		row4.createCell(9).setCellValue("");
		
		Row row5 = nrRequestSheet.createRow(5);
		row5.createCell(0).setCellValue("Cost Centre");
		row5.createCell(1).setCellValue("");
		row5.createCell(6).setCellValue("Quote Prive");
		row5.createCell(7).setCellValue("");
		
		Row row6 = nrRequestSheet.createRow(6);
		row6.createCell(0).setCellValue("Request Reason");
		row6.createCell(1).setCellValue("");
		
		Row row7 = nrRequestSheet.createRow(7);
		row7.createCell(0).setCellValue("Duration");
		row7.createCell(1).setCellValue("");
		row7.createCell(2).setCellValue("");
		row7.createCell(3).setCellValue("");
		row7.createCell(4).setCellValue("Product Code");
		row7.createCell(5).setCellValue("");
		row7.createCell(6).setCellValue("");
		
		
		Sheet nrApproversheet = wb.createSheet("Non Approver");
		
		
		Sheet nrLogsheet = wb.createSheet("Non Log");
		
		response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-disposition", "attachment;filename=" + "hello.xlsx");
        response.flushBuffer();
        OutputStream outputStream = response.getOutputStream();
        wb.write(response.getOutputStream());
        wb.close();
        outputStream.flush();
        outputStream.close();
        
		return String.format("Hello %s!", name);
	}
}