package support;

public class ExcelDataMigrationException extends RuntimeException {

	private static final long serialVersionUID = -4362303779150904109L;
	
	public ExcelDataMigrationException() {
		super();
	}
	
	public ExcelDataMigrationException(String message) {
		super(message);
	}
	
	public ExcelDataMigrationException(String message, Throwable throwable) {
		super(message, throwable);
	}
}
