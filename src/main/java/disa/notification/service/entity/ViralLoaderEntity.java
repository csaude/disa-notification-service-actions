package disa.notification.service.entity;

import java.time.LocalDate;
import java.time.LocalDateTime;

import javax.persistence.Basic;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.EnumType;
import javax.persistence.Enumerated;
import javax.persistence.Id;
import javax.persistence.Table;

import disa.notification.service.enums.NotProcessingCause;
import disa.notification.service.enums.ViralLoadStatus;
import lombok.Data;

@Entity(name = ViralLoaderEntity.TABLE_NAME)
@Table(name = ViralLoaderEntity.TABLE_NAME)
@Data
public class ViralLoaderEntity {

    @Id
    private Long id;

    public static final String TABLE_NAME = "VlData";
    @Column(name = "UNIQUEID")
    private String nid;

    @Column(name = "RequestID")
    private String requestId;

    @Column(name = "ReferringRequestID")
    private String referringRequestID;

    @Column(name = "FIRSTNAME")
    private String firstName;

    @Column(name = "SURNAME")
    private String lastName;

    @Column(name = "HL7SexCode")
    private String gender;

    @Column(name = "DOB")
    private LocalDate dateOfBirth;

    @Column(name = "LOCATION")
    private String location;

    @Column(name = "RequestingFacilityCode")
    private String healthFacilityLabCode;

    @Column(name = "RequestingFacilityName")
    private String requestingFacilityName;

    @Column(name = "AttendingDoctor")
    private String nameOfTechnicianRequestingTest;

    @Column(name = "WARD")
    private String encounter;

    @Column(name = "Pregnant")
    private String pregnant;

    @Column(name = "BreastFeeding")
    private String breastFeeding;

    @Column(name = "ReasonForTest")
    private String reasonForTest;

    @Column(name = "SpecimenDatetime")
    private LocalDateTime harvestDate;

    @Column(name = "TypeOfSampleCollection")
    private String harvestType;

    @Column(name = "ReceivedDateTime")
    private LocalDateTime dateOfSampleReceive;

    @Column(name = "LIMSRejectionDesc")
    private String rejectedReason;

    @Column(name = "AnalysisDateTime")
    private LocalDateTime processingDate;

    @Column(name = "LIMSSpecimenSourceCode")
    private String sampleType;

    @Column(name = "HIVVL_ViralLoadCAPCTM")
    private String viralLoadResultCopies;

    @Column(name = "HIVVL_VRLogValue")
    private String viralLoadResultLog;

    @Column(name = "ViralLoadResultCategory")
    private String viralLoadResultQualitative;

    @Column(name = "AuthorisedDateTime")
    private LocalDateTime viralLoadResultDate;

    @Column(name = "ClinicalInfo")
    private String labComments;

    @Basic(optional = false)
    @Column(name = "VIRAL_LOAD_STATUS", columnDefinition = "enum('PENDING','PROCESSED','NOT_PROCESSED')")
    @Enumerated(EnumType.STRING)
    private ViralLoadStatus viralLoadStatus;

    @Column(name = "HIVVL_ViralLoadResult")
    private String hivViralLoadResult;

    @Enumerated(EnumType.STRING)
    @Column(name = "NOT_PROCESSING_CAUSE")
    private NotProcessingCause notProcessingCause;

    @Column(name = "CREATED_AT")
    private LocalDateTime createdAt;

}
