package disa.notification.service.utils;

public interface XLSColumnConstants {

    enum ResultsReceivedByNid {
        REQUEST_ID,
        TYPE_OF_RESULT,
        NID,
        PROVINCE,
        DISTRICT,
        HEALTH_FACILITY_CODE,
        HEALTH_FACILITY_NAME,
        CREATED_AT,
        UPDATED_AT,
        VIRAL_RESULT_STATUS,
        NOT_PROCESSING_CAUSE,
        OBS;
    }

    enum ResultsByHFSummary {
        PROVINCE,
        DISTRICT,
        HEALTH_FACILITY_CODE,
        HEALTH_FACILITY_NAME,
        TYPE_OF_RESULT,
        TOTAL_RECEIVED,
        TOTAL_PROCESSED,
        TOTAL_PENDING,
        NOT_PROCESSED_INVALID_RESULT,
        NOT_PROCESSED_NID_NOT_FOUND,
        NOT_PROCESSED_DUPLICATED_NID,
        NOT_PROCESSED_DUPLICATED_REQUEST_ID;
    }

    enum ResultsPendingByUs {
        PROVINCE,
        DISTRICT,
        US_CODE,
        US_NAME,
        TOTAL_PENDING,
        LAST_SYNC;
    }

    enum ResultsByDistrictSummary {
        PROVINCE,
        DISTRICT,
        TYPE_OF_RESULT,
        TOTAL_PROCESSED,
        PERCENTAGE_PROCESSED,
        TOTAL_PENDING,
        PERCENTAGE_PENDING,
        NOT_PROCESSED_INVALID_RESULT,
        PERCENTAGE_NOT_PROCESSED_INVALID_RESULT,
        NOT_PROCESSED_NID_NOT_FOUND,
        PERCENTAGE_NOT_PROCESSED_NID_NOT_FOUND,
        NOT_PROCESSED_DUPLICATED_NID,
        PERCENTAGE_NOT_PROCESSED_DUPLICATED_NID,
        NOT_PROCESSED_DUPLICATED_REQUEST_ID,
        PERCENTAGE_NOT_PROCESSED_DUPLICATED_REQUEST_ID,
        TOTAL_RECEIVED;
    }

    enum ResultsPendingByNid {
        REQUEST_ID,
        NID,
        PROVINCE,
        DISTRICT,
        HEALTH_FACILITY_CODE,
        HEALTH_FACILITY_NAME,
        SENT_DATE,
        STATUS;
    }
}
