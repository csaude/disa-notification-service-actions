package disa.notification.service.utils;

import disa.notification.service.service.interfaces.LabResultSummary;
import lombok.Builder;
import lombok.Getter;

@Builder
@Getter
public class VlResultSummaryImpl implements LabResultSummary {
	private String requestingProvinceName;
    private String requestingDistrictName;
    private String healthFacilityLabCode;
    private String typeOfResult;
    private String facilityName;
    private int totalReceived;
    private int totalPending;
    private int processed;
    private int notProcessedInvalidResult;
    private int notProcessedNidNotFount;
    private int notProcessedDuplicateNid;
    private int notProcessedDuplicatedRequestId;
}
