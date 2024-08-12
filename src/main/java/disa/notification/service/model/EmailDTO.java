package disa.notification.service.model;

public class EmailDTO {

	private String[] to;
	private String subject;
	private String body;
	private String module;
	private String startDate;
	private String endDate;
	private String repoLink;
	private Boolean resultFlag;
	
	public String[] getTo() {
		return to;
	}
	public void setTo(String[] mailList) {
		this.to = mailList;
	}
	public String getSubject() {
		return subject;
	}
	public void setSubject(String subject) {
		this.subject = subject;
	}
	public String getBody() {
		return body;
	}
	public void setBody(String body) {
		this.body = body;
	}
	public String getModule() {
		return module;
	}
	public void setModule(String module) {
		this.module = module;
	}
	public String getStartDate() {
		return startDate;
	}
	public void setStartDate(String startDate) {
		this.startDate = startDate;
	}
	public String getEndDate() {
		return endDate;
	}
	public void setEndDate(String endDate) {
		this.endDate = endDate;
	}
	public String getRepoLink() {
		return repoLink;
	}
	public void setRepoLink(String repoLink) {
		this.repoLink = repoLink;
	}
	public void setResultFlag(Boolean resultFlag) {
		this.resultFlag = resultFlag;
	}
	public Boolean getResultFlag() {
		return resultFlag;
	}
}
