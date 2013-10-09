
public class layersClassObject {

	private String modelName;
	private String topic;
	private String className;
	private String rowIndex;
	private boolean haveSame = false;
	
	public layersClassObject(String modelName, String topic, String className, String rowIndex) {

		this.modelName = modelName;
		this.topic = topic;
		this.className = className;
		this.rowIndex = rowIndex;
	}
	
	public boolean isHaveSame() {
		return haveSame;
	}

	public void setHaveSame(boolean haveSame) {
		this.haveSame = haveSame;
	}

	public String getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(String rowIndex) {
		this.rowIndex = rowIndex;
	}



	public String getTopic() {
		return topic;
	}

	public void setTopic(String topic) {
		this.topic = topic;
	}

	public String getClassName() {
		return className;
	}

	public void setClassName(String className) {
		this.className = className;
	}

	public String getModelName() {
		return modelName;
	}

	public void setModelName(String modelName) {
		this.modelName = modelName;
	}	
}
