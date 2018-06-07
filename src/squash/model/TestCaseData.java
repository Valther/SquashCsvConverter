package squash.model;

import java.util.ArrayList;
import java.util.List;

public class TestCaseData {
	String id;
	String description;
	String preRequisite;
	String nature;
	String dateModification;
	String createdBy;
	String hierarchy;
	String label;
	String weight;
	List<String> actions, expectedResults;

	public TestCaseData() {
		actions = new ArrayList<String>();
		expectedResults = new ArrayList<String>();
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}
	
	public String getHierarchy() {
		return hierarchy;
	}

	public void setHierarchy(String hierarchy) {
		this.hierarchy = hierarchy;
	}

	public String getLabel() {
		return label;
	}

	public void setLabel(String label) {
		this.label = label;
	}

	public List<String> getActions() {
		return actions;
	}

	public void addAction(String action) {
		this.actions.add(action);
	}

	public List<String> getExpectedResults() {
		return expectedResults;
	}

	public void addExpectedResults(String result) {
		this.expectedResults.add(result);
	}

	public String getWeight() {
		return weight;
	}

	public void setWeight(String weight) {
		this.weight = weight;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getPreRequisite() {
		return preRequisite;
	}

	public void setPreRequisite(String preRequisite) {
		this.preRequisite = preRequisite;
	}

	public String getNature() {
		return nature;
	}

	public void setNature(String nature) {
		this.nature = nature;
	}

	public String getDateModification() {
		return dateModification;
	}

	public void setDateModification(String dateModification) {
		this.dateModification = dateModification;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	@Override
	public String toString() {
		return "TestCaseData [id=" + id + ", description=" + description + ", preRequisite=" + preRequisite
				+ ", nature=" + nature + ", dateModification=" + dateModification + ", createdBy=" + createdBy
				+ ", hierarchy=" + hierarchy + ", label=" + label + ", actions=" + actions + ", expectedResults="
				+ expectedResults + "]";
	}

}
