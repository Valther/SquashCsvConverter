package squash.model;

public class LineReportSquash {
	private String path, id, reference, label, weight, nature, type, status, description, pre_requisite, created_on, created_by,
			last_modified_on, last_modified_by, action, expected_result;

	public LineReportSquash(String path, String id, String reference, String label, String weight, String nature, String type,
			String status, String description, String pre_requisite, String created_on, String created_by,
			String last_modified_on, String last_modified_by, String action, String expected_result) {
		this.path = path;
		this.id = id;
		this.reference = reference;
		this.label = label;
		this.weight = weight;
		this.nature = nature;
		this.type = type;
		this.status = status;
		this.description = description;
		this.pre_requisite = pre_requisite;
		this.created_on = created_on;
		this.created_by = created_by;
		this.last_modified_on = last_modified_on;
		this.last_modified_by = last_modified_by;
		this.action = action;
		this.expected_result = expected_result;
	}

	public String getPath() {
		return path;
	}

	public String getReference() {
		return reference;
	}

	public String getId() {
		return id;
	}

	public String getLabel() {
		return label;
	}

	public String getWeight() {
		return weight;
	}

	public String getNature() {
		return nature;
	}

	public String getType() {
		return type;
	}

	public String getStatus() {
		return status;
	}

	public String getDescription() {
		return description;
	}

	public String getPre_requisite() {
		return pre_requisite;
	}

	public String getCreated_on() {
		return created_on;
	}

	public String getCreated_by() {
		return created_by;
	}

	public String getLast_modified_on() {
		return last_modified_on;
	}

	public String getLast_modified_by() {
		return last_modified_by;
	}

	public String getAction() {
		return action;
	}

	public String getExpected_result() {
		return expected_result;
	}

	@Override
	public String toString() {
		return "LineExcelSquash [path=" + path + ", id=" + id + ", reference=" + reference + ", label=" + label
				+ ", weight=" + weight + ", nature=" + nature + ", type=" + type + ", status=" + status
				+ ", description=" + description + ", pre_requisite=" + pre_requisite + ", created_on=" + created_on
				+ ", created_by=" + created_by + ", last_modified_on=" + last_modified_on + ", last_modified_by="
				+ last_modified_by + ", action=" + action + ", expected_result=" + expected_result + "]";
	}

	
}
