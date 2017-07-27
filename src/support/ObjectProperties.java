package support;

public class ObjectProperties {

	private String attributeName;
	private Class<?> attributeType;
	private String columnName;
	
	public ObjectProperties(String attributeName, Class<?> attributeType, String columnName) {
		this.attributeName = attributeName;
		this.attributeType = attributeType;
		this.columnName = columnName;
	}
	
	public ObjectProperties(String attributeName, Class<?> attributeType) {
		this(attributeName, attributeType, null);
	}
	
	public ObjectProperties(String attributeName, String columnName) {
		this(attributeName, null, columnName);
	}
	
	public ObjectProperties(String attributeName) {
		this(attributeName, null, null);
	}
	
	public String getAttributeName() {
		return attributeName;
	}
	public void setAttributeName(String attributeName) {
		this.attributeName = attributeName;
	}
	public Class<?> getAttributeType() {
		return attributeType;
	}
	public void setAttributeType(Class<?> attributeType) {
		this.attributeType = attributeType;
	}
	public String getColumnName() {
		return columnName == null ? attributeName : columnName;
	}
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}
}
