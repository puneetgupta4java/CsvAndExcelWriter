package com.example.demo;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.supercsv.cellprocessor.ParseInt;
import org.supercsv.cellprocessor.constraint.NotNull;
import org.supercsv.cellprocessor.ift.CellProcessor;
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;

import com.opencsv.bean.ColumnPositionMappingStrategy;
import com.opencsv.bean.StatefulBeanToCsv;
import com.opencsv.bean.StatefulBeanToCsvBuilder;

@RestController
@RequestMapping("/add")
public class EmployeeController {
	List<Employee> EmployeeList = new ArrayList<Employee>();

	public EmployeeController() {
		Employee emp1 = new Employee(1, "puneet", 24, new Address("abc1", 101));
		Employee emp2 = new Employee(2, "Aman", 24, new Address("abc2", 102));
		Employee emp3 = new Employee(3, "Suvradip", 26, new Address("abc3", 103));
		Employee emp4 = new Employee(4, "Riya", 22, new Address("abc4", 104));
		Employee emp5 = new Employee(5, "Prakash", 29, new Address("abc5", 105));
		EmployeeList.add(emp1);
		EmployeeList.add(emp2);
		EmployeeList.add(emp3);
		EmployeeList.add(emp4);
		EmployeeList.add(emp5);
	}

	@GetMapping("/opencsv")
	public String addEmployees() {
		final String CSV_LOCATION = "Employeesopencsv.csv ";

		try {

			FileWriter writer = new FileWriter(CSV_LOCATION);
			ColumnPositionMappingStrategy<Employee> mappingStrategy = new ColumnPositionMappingStrategy<Employee>();
			mappingStrategy.setType(Employee.class);
			String[] columns = new String[] { "Id", "Name", "Age", "Street", "Pincode" };
			mappingStrategy.setColumnMapping(columns);
			StatefulBeanToCsvBuilder<Employee> builder = new StatefulBeanToCsvBuilder<Employee>(writer);
			StatefulBeanToCsv<Employee> beanWriter = builder.withMappingStrategy(mappingStrategy).build();
			beanWriter.write(EmployeeList);
			writer.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "added using open csv success";
	}

	@GetMapping("/supercsv")
	public String addEmployeesSupercsv() {

		ICsvBeanWriter beanWriter = null;
		final CellProcessor[] processors = new CellProcessor[] { new NotNull(new ParseInt()), new NotNull(),
				new NotNull(new ParseInt()), new NotNull() };
		try {
			beanWriter = new CsvBeanWriter(new FileWriter("Employeessupercsv.csv"), CsvPreference.STANDARD_PREFERENCE);
			final String[] header = new String[] { "Id", "Name", "Age", "Address" };
			beanWriter.writeHeader(header);
			for (Employee c : EmployeeList) {
				beanWriter.write(c, header, processors);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (beanWriter != null) {
				try {
					beanWriter.close();
				} catch (IOException ex) {
					System.err.println("Error closing the writer: " + ex);
				}
			}
		}

		return "added using super csv  success";
	}

	@GetMapping("/excel")
	public String addEmployeeExcel() {

		String[] COLUMNs = { "Id", "Name", "Age", "Street", "Pincode" };
		try (Workbook workbook = new XSSFWorkbook(); OutputStream fileOut = new FileOutputStream("employees.xlsx");) {
			CreationHelper createHelper = workbook.getCreationHelper();

			Sheet sheet = workbook.createSheet("Employees");

			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setColor(IndexedColors.BLUE.getIndex());

			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);

			// Row for Header
			Row headerRow = sheet.createRow(0);

			// Header
			for (int col = 0; col < COLUMNs.length; col++) {
				Cell cell = headerRow.createCell(col);
				cell.setCellValue(COLUMNs[col]);
				cell.setCellStyle(headerCellStyle);
			}

			// CellStyle for Age
			CellStyle ageCellStyle = workbook.createCellStyle();
			ageCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));
			ageCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			ageCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			int rowIdx = 1;
			for (Employee employee : EmployeeList) {
				Row row = sheet.createRow(rowIdx++);

				row.createCell(0).setCellValue(employee.getId());
				row.createCell(1).setCellValue(employee.getName());
				Cell ageCell = row.createCell(2);
				ageCell.setCellValue(employee.getAge());
				if(employee.getAge() > 25)
					ageCell.setCellStyle(ageCellStyle);
				row.createCell(3).setCellValue(employee.getAddress().getStreet());
				row.createCell(4).setCellValue(employee.getAddress().getPincode());
			}

			workbook.write(fileOut);
		} catch (Exception e) {
			System.out.println(e.getLocalizedMessage());
		}
		return "added to excel success";
	}
}
