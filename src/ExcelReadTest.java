import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import support.ExcelDataMigration;
import support.ObjectProperties;

public class ExcelReadTest {
	
	private static final String filePath = "C:\\development\\project-temp\\java-ExcelMigration\\temporary\\Book1.xlsx";
	
	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException, InstantiationException, IllegalAccessException, NoSuchMethodException, SecurityException, IllegalArgumentException, InvocationTargetException, NoSuchFieldException {
		
		long convertMapStartTime = System.currentTimeMillis();
		
		// map test
		ExcelDataMigration excelDataMigration1 = new ExcelDataMigration(filePath);
		List<ObjectProperties> propertiesList1 = new ArrayList<>();
		propertiesList1.add(new ObjectProperties("column1", "컬럼1"));
		propertiesList1.add(new ObjectProperties("column2", "컬럼2"));
		propertiesList1.add(new ObjectProperties("column3", "컬럼3"));
		propertiesList1.add(new ObjectProperties("column4", "컬럼4"));
		propertiesList1.add(new ObjectProperties("column5", "컬럼5"));
		propertiesList1.add(new ObjectProperties("column6", "컬럼6"));
		propertiesList1.add(new ObjectProperties("column7", "컬럼7"));
		
		List<Map<String, Object>> resultList1 = (List<Map<String, Object>>) excelDataMigration1.convertMapList(propertiesList1);
		
		long convertMapStopTime1 = System.currentTimeMillis();
		long convertMapExcuteTime = convertMapStopTime1 - convertMapStartTime;
		
		for (Map<String, Object> item : resultList1) {
			System.out.println(item.toString());
		}
		
		long convertMapStopTime2 = System.currentTimeMillis();
		long convertMapPrintTime = convertMapStopTime2 - convertMapStopTime1;
		
		System.out.println("======================================");
		System.out.println("======================================");
		
		long convertObjectStartTime = System.currentTimeMillis();
		
		// object test
		ExcelDataMigration excelDataMigration2 = new ExcelDataMigration(filePath);
		List<ObjectProperties> propertiesList2 = new ArrayList<>();
		propertiesList2.add(new ObjectProperties("column1", "컬럼1"));
		propertiesList2.add(new ObjectProperties("column2", "컬럼2"));
		propertiesList2.add(new ObjectProperties("column3", "컬럼3"));
		propertiesList2.add(new ObjectProperties("column4", "컬럼4"));
		propertiesList2.add(new ObjectProperties("column5", "컬럼5"));
		propertiesList2.add(new ObjectProperties("column6", "컬럼6"));
		propertiesList2.add(new ObjectProperties("column7", "컬럼7"));
		
		long convertObjectStopTime1 = System.currentTimeMillis();
		long convertObjectExcuteTime = convertObjectStopTime1 - convertObjectStartTime;
		
		List<ExcelObject> resultList2 = (List<ExcelObject>) excelDataMigration2.convertObjectList(ExcelObject.class, propertiesList2);
		for (ExcelObject item : resultList2) {
			System.out.println(item.toString());
		}
		
		long convertObjectStopTime2 = System.currentTimeMillis();
		long convertObjectPrintTime = convertObjectStopTime2 - convertObjectStopTime1;
		
		
		System.out.println(" convertMapExcuteTime > " + convertMapExcuteTime);
		System.out.println(" convertMapPrintTime  > " + convertMapPrintTime);
		
		System.out.println(" convertObjectExcuteTime > " + convertObjectExcuteTime);
		System.out.println(" convertObjectPrintTime  > " + convertObjectPrintTime);
	}
}
