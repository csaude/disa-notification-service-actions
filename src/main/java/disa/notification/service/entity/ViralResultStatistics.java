package disa.notification.service.entity;

import disa.notification.service.service.interfaces.LabResultSummary;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ViralResultStatistics {
	
	private String typeOfResult; 
    private int total;
    private int processed;
    private int pending;
    private int noProcessedInvalidResult;
    private int noProcessedNidNotFound;
    private int notProcessedDuplicateNid;
    private int notProcessedDuplicatedReqId;

    public void accumulate(LabResultSummary summary) {
    	typeOfResult = summary.getTypeOfResult();
        total += summary.getTotalReceived();
        processed += summary.getProcessed();
        pending += summary.getTotalPending();
        noProcessedInvalidResult += summary.getNotProcessedInvalidResult();
        noProcessedNidNotFound += summary.getNotProcessedNidNotFount();
        notProcessedDuplicateNid += summary.getNotProcessedDuplicateNid();
        notProcessedDuplicatedReqId += summary.getNotProcessedDuplicatedRequestId();
    }

    public void accumulate(ViralResultStatistics stats) {
        total += stats.getTotal();
        processed += stats.getProcessed();
        pending += stats.getPending();
        noProcessedInvalidResult += stats.getNoProcessedInvalidResult();
        noProcessedNidNotFound += stats.getNoProcessedNidNotFound();
        notProcessedDuplicateNid += stats.getNotProcessedDuplicateNid();
        notProcessedDuplicatedReqId += stats.getNotProcessedDuplicatedReqId();
    }

    public ViralResultStatistics combine(ViralResultStatistics other) {
        this.accumulate(other);
        return this;
    }

    public double getProcessedPercentage() {
        return total == 0 ? 0.0 : (double) processed / total;
    }

    public double getPendingPercentage() {
        return total == 0 ? 0.0 : (double) pending / total;
    }

    public double getNoProcessedNoResultPercentage() {
        return total == 0 ? 0.0 : (double) noProcessedInvalidResult / total;
    }

    public double getNoProcessedNidNotFoundPercentage() {
        return total == 0 ? 0.0 : (double) noProcessedNidNotFound / total;
    }

    public double getNotProcessedDuplicateNidPercentage() {
        return total == 0 ? 0.0 : (double) notProcessedDuplicateNid / total;
    }

    public double getNotProcessedDuplicatedReqIdPercentage() {
        return total == 0 ? 0.0 : (double) notProcessedDuplicatedReqId / total;
    }

	public String getTypeOfResult() {
		return typeOfResult;
	}
}
