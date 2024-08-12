package disa.notification.service.service.interfaces;

public interface LabResultSummary {
	String getRequestingProvinceName();
    String getRequestingDistrictName();
    String getHealthFacilityLabCode();
    String getTypeOfResult();
    String getFacilityName();
    int getTotalReceived();
    int getTotalPending();
    int getProcessed();
    int getNotProcessedInvalidResult();
    int getNotProcessedNidNotFount();
    int getNotProcessedDuplicateNid();
    int getNotProcessedDuplicatedRequestId();
}