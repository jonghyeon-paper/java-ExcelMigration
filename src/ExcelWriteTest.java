import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import support.ExcelDataMigration;
import support.ObjectProperties;

public class ExcelWriteTest {
	
	private final static String filePath = "C:\\development\\project-temp\\java-ExcelMigration\\temporary\\Book2.xlsx";

	public static void main(String[] args) throws IOException {
		List<ObjectProperties> propertiesList1 = new ArrayList<>();
		propertiesList1.add(new ObjectProperties("column1", "컬럼1"));
		propertiesList1.add(new ObjectProperties("column2", "컬럼2"));
		propertiesList1.add(new ObjectProperties("column3", "컬럼3"));
		propertiesList1.add(new ObjectProperties("column4", "컬럼4"));
		propertiesList1.add(new ObjectProperties("column5", "컬럼5"));
		propertiesList1.add(new ObjectProperties("column6", "컬럼6"));
		propertiesList1.add(new ObjectProperties("column7", "컬럼7"));
		
		List<ExcelObject> list = new ArrayList<>();
		list.add(new ExcelObject(1, "일", "하나",  "첫번쨰", "ㄱ", "1", "a"));
		list.add(new ExcelObject(2, "이",  "둘",  "두번쨰", "ㄴ", "2", "b"));
		list.add(new ExcelObject(3, "삼",  "셋",  "세번쨰", "ㄷ", "3", "c"));
		list.add(new ExcelObject(4, "사",  "넷",  "네번쨰", "ㄹ", "4", "d"));
		list.add(new ExcelObject(5, "오", "다섯", "다섯번쨰", "ㅁ", "5", "e"));
		list.add(new ExcelObject(6, "육", "여섯", "여섯번쨰", "ㅂ", "6", "f"));
		list.add(new ExcelObject(7, "칠", "일곱", "일곱번쨰", "ㅅ", "7", "g"));
		list.add(new ExcelObject(8, "팔", "여덟", "여덟번쨰", "ㅇ", "8", "h"));
		list.add(new ExcelObject(9, "구", "아홉", "아홉번쨰", "ㅈ", "9", "i"));
		
		ExcelDataMigration excelDataMigration = new ExcelDataMigration();
		XSSFWorkbook workbook = excelDataMigration.createXlsxFromObjectList(list, propertiesList1);
		
		File writrFile = new File(filePath);
		FileOutputStream fos = new FileOutputStream(writrFile);
		workbook.write(fos);
	}

}
