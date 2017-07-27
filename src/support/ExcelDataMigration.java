package support;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataMigration {
	
	private InputStream inputStream;
	private Boolean existTitleRow = true;
	private Boolean createTitleRow = true;
	
	public ExcelDataMigration() {
	}
	
	public ExcelDataMigration(InputStream inputStream) {
		this.inputStream = inputStream;
	}
	
	public ExcelDataMigration(File file) throws FileNotFoundException {
		this(new FileInputStream(file));
	}
	
	public ExcelDataMigration(String path) throws FileNotFoundException {
		this(new File(path));
	}
	
	public Boolean isExistTitleRow() {
		return existTitleRow;
	}

	public void setExistTitleRow(Boolean existTitleRow) {
		this.existTitleRow = existTitleRow;
	}

	public Boolean isCreateTitleRow() {
		return createTitleRow;
	}

	public void setCreateTitleRow(Boolean createTitleRow) {
		this.createTitleRow = createTitleRow;
	}

	public List<? extends Object> convertObjectList(Class<?> objectClass, List<ObjectProperties> objectPropertiesList) {
		List<Object> resultList = new ArrayList<>();
		List<ObjectProperties> sortedObjectPropertiesList = new ArrayList<>();
		
		try {
			// Step1. sheet check
			@SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			int sheets = workbook.getNumberOfSheets();
			for (int i = 0; i < sheets; i++) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				// Step2. row check
				int rows = sheet.getPhysicalNumberOfRows();
				for (int j = 0; j < rows; j++) {
					XSSFRow row = sheet.getRow(j);
					
					// Step3. cell check
					Object newInstance = objectClass.newInstance();
					int cells = row.getPhysicalNumberOfCells();
					for (int k = 0; k < cells; k++) {
						XSSFCell cell = row.getCell(k);
						
						if (j == 0 && existTitleRow) {
							// 컬럼 설정
							String excelColumnName = cell.getStringCellValue();
							for (ObjectProperties item : objectPropertiesList) {
								String propertyColumnName = item.getColumnName() == null ? item.getAttributeName() : item.getColumnName();
								if (!excelColumnName.equals(propertyColumnName)) {
									continue;
								}
								sortedObjectPropertiesList.add(item);
								break;
							}
						} else {
							// 데이터 설정
							Object value = null;
							switch (cell.getCellTypeEnum()) {
							case ERROR :
								value = "ERROR!";
								break;
							case BOOLEAN :
								value = cell.getBooleanCellValue() ? "true" : "false";
								break;
							case BLANK :
								value = "";
								break;
							case NUMERIC :
								value = String.valueOf(cell.getNumericCellValue());
								break;
							case STRING :
							default :
								value = cell.getStringCellValue();
								break;
							}
							
							ObjectProperties columnProperty = existTitleRow ? sortedObjectPropertiesList.get(k) : objectPropertiesList.get(k);
							setObjectValue(newInstance, columnProperty, value);
						}
					}
					
					if (j == 0 && existTitleRow) {
						continue;
					}
					resultList.add(newInstance);
				}
			}
		} catch (IOException e) {
			//e.printStackTrace();
			throw new ExcelDataMigrationException("stream exception");
		} catch (InstantiationException | IllegalAccessException e) {
			e.printStackTrace();
			throw new ExcelDataMigrationException("generate instance exception");
		} catch (NoSuchFieldException | NoSuchMethodException | IllegalArgumentException | InvocationTargetException e) {
			//e.printStackTrace();
			throw new ExcelDataMigrationException("method reflection exception");
		}
		return resultList;
	}
	
	private void setObjectValue(Object newInstance, ObjectProperties property, Object value)
			throws NoSuchFieldException, NoSuchMethodException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		if (property.getAttributeType() == null) {
			Field field = newInstance.getClass().getDeclaredField(property.getAttributeName());
			property.setAttributeType(field.getType());
		}
		
		Pattern pattern = Pattern.compile("^[\\w]{1}");
		Matcher matcher = null;
		
		matcher = pattern.matcher(property.getAttributeName());
		if (matcher.find()) {
			String setMethodName = matcher.replaceAll("set" + matcher.group().toUpperCase());
			Method setMethod = newInstance.getClass().getDeclaredMethod(setMethodName, new Class[] {property.getAttributeType()});
			if (Number.class.isAssignableFrom(property.getAttributeType())) {
				Double doubleValue = Double.valueOf(String.valueOf(value));
				Integer finalValue = doubleValue.intValue();
				setMethod.invoke(newInstance, finalValue);
			}
			if (String.class.isAssignableFrom(property.getAttributeType())) {
				String finalValue = String.valueOf(value);
				setMethod.invoke(newInstance, finalValue);
			}
		}
	}
	
	public List<? extends Map<String, Object>> convertMapList(List<ObjectProperties> objectPropertiesList) {
		List<Map<String, Object>> resultList = new ArrayList<>();
		List<ObjectProperties> sortedObjectPropertiesList = new ArrayList<>();
		
		try {
			// Step1. sheet check
			@SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			int sheets = workbook.getNumberOfSheets();
			for (int i = 0; i < sheets; i++) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				// Step2. row check
				int rows = sheet.getPhysicalNumberOfRows();
				for (int j = 0; j < rows; j++) {
					XSSFRow row = sheet.getRow(j);
					
					// Step3. cell check
					Map<String, Object> map = new HashMap<>();
					int cells = row.getPhysicalNumberOfCells();
					for (int k = 0; k < cells; k++) {
						XSSFCell cell = row.getCell(k);
						
						if (j == 0 && existTitleRow) {
							// 컬럼 설정
							String excelColumnName = cell.getStringCellValue();
							for (ObjectProperties item : objectPropertiesList) {
								String propertyColumnName = item.getColumnName() == null ? item.getAttributeName() : item.getColumnName();
								if (!excelColumnName.equals(propertyColumnName)) {
									continue;
								}
								sortedObjectPropertiesList.add(item);
								break;
							}
						} else {
							// 데이터 설정
							Object value = null;
							switch (cell.getCellTypeEnum()) {
							case ERROR :
								value = "ERROR!";
								break;
							case BOOLEAN :
								value = cell.getBooleanCellValue() ? "true" : "false";
								break;
							case BLANK :
								value = "";
								break;
							case NUMERIC :
								value = String.valueOf(cell.getNumericCellValue());
								break;
							case STRING :
							default :
								value = cell.getStringCellValue();
								break;
							}
							
							ObjectProperties columnProperty = existTitleRow ? sortedObjectPropertiesList.get(k) : objectPropertiesList.get(k);
							map.put(columnProperty.getAttributeName(), value);
						}
					}
					
					if (j == 0 && existTitleRow) {
						continue;
					}
					resultList.add(map);
				}
			}
		} catch (IOException e) {
			//e.printStackTrace();
			throw new ExcelDataMigrationException("stream exception");
		}
		return resultList;
	}
	
	public XSSFWorkbook createXlsxFromObjectList(List<? extends Object> dataList, List<ObjectProperties> objectPropertiesList) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet1");
		
		try {
			int rowIndex = 0;
			if (createTitleRow) {
				XSSFRow titleRow = sheet.createRow(rowIndex);
				
				for (int cellIndex = 0; cellIndex < objectPropertiesList.size(); cellIndex++) {
					ObjectProperties property = objectPropertiesList.get(cellIndex);
					if (property.getColumnName() == null) {
						property.setColumnName(property.getAttributeName());
					}
					
					XSSFCell cell = titleRow.createCell(cellIndex);
					cell.setCellValue(objectPropertiesList.get(cellIndex).getColumnName());
				}
				rowIndex++;
			}
			
			for (Object item : dataList) {
				XSSFRow dataRow = sheet.createRow(rowIndex);
				
				for (int cellIndex = 0; cellIndex < objectPropertiesList.size(); cellIndex++) {
					ObjectProperties property = objectPropertiesList.get(cellIndex);
					XSSFCell cell = dataRow.createCell(cellIndex);
					
					setCellValue(cell, property, item);
				}
				rowIndex++;
			}
		} catch (NoSuchMethodException | SecurityException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
			//e.printStackTrace();
			throw new ExcelDataMigrationException("method reflection exception");
		}
		return workbook;
	}
	
	private void setCellValue(XSSFCell newCell, ObjectProperties property, Object data) throws NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Pattern pattern = Pattern.compile("^[\\w]{1}");
		Matcher matcher = null;
		
		matcher = pattern.matcher(property.getAttributeName());
		if (matcher.find()) {
			String getMethodName = matcher.replaceAll("get" + matcher.group().toUpperCase());
			Method getMethod = data.getClass().getDeclaredMethod(getMethodName);
			
			Object value = getMethod.invoke(data);
			if (Number.class.isAssignableFrom(value.getClass())) {
				Integer finalValue = Integer.parseInt(String.valueOf(value));
				newCell.setCellValue(finalValue);
			}
			if (String.class.isAssignableFrom(value.getClass())) {
				String finalValue = String.valueOf(value);
				newCell.setCellValue(finalValue);
			}
		}
	}
	
	public XSSFWorkbook createXlsxFromMapList(List<? extends Map<String, Object>> dataList, List<ObjectProperties> objectPropertiesList) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet1");
		
		int rowIndex = 0;
		if (createTitleRow) {
			XSSFRow titleRow = sheet.createRow(rowIndex);
			
			for (int cellIndex = 0; cellIndex < objectPropertiesList.size(); cellIndex++) {
				ObjectProperties property = objectPropertiesList.get(cellIndex);
				if (property.getColumnName() == null) {
					property.setColumnName(property.getAttributeName());
				}
				
				XSSFCell cell = titleRow.createCell(cellIndex);
				cell.setCellValue(objectPropertiesList.get(cellIndex).getColumnName());
			}
			rowIndex++;
		}
		
		for (Map<String, Object> item : dataList) {
			XSSFRow dataRow = sheet.createRow(rowIndex);
			
			for (int cellIndex = 0; cellIndex < objectPropertiesList.size(); cellIndex++) {
				ObjectProperties property = objectPropertiesList.get(cellIndex);
				XSSFCell cell = dataRow.createCell(cellIndex);
				
				Object value = item.get(property.getAttributeName());
				if (Number.class.isAssignableFrom(value.getClass())) {
					Integer finalValue = Integer.parseInt(String.valueOf(value));
					cell.setCellValue(finalValue);
				}
				if (String.class.isAssignableFrom(value.getClass())) {
					String finalValue = String.valueOf(value);
					cell.setCellValue(finalValue);
				}
			}
			rowIndex++;
		}
		return workbook;
	}
}
