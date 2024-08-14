package disa.notification.service.utils;

import static java.util.stream.Collectors.groupingBy;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.context.MessageSource;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;

import disa.notification.service.entity.ViralResultStatistics;
import disa.notification.service.enums.ViralLoadStatus;
import disa.notification.service.service.interfaces.LabResultSummary;
import disa.notification.service.service.interfaces.LabResults;
import disa.notification.service.service.interfaces.PendingHealthFacilitySummary;

public class SyncReport implements XLSColumnConstants {

    private static final int VARIABLES_SHEET = 0;
    // Dictionary sheet = 1
    private static final int RECEIVED_BY_DISTRICT_SHEET = 2;
    private static final int RECEIVED_BY_US_SHEET = 3;
    private static final int RECEIVED_BY_NID_SHEET = 4;
    private static final int PENDING_BY_US_SHEET = 5;
    private static final int PENDING_BY_NID_SHEET = 6;

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");

    private MessageSource messageSource;

    private DateInterval reportDateInterval;

    public SyncReport(MessageSource messageSource, DateInterval reportDateInterval) {
        this.messageSource = messageSource;
        this.reportDateInterval = reportDateInterval;
    }

    public ByteArrayResource getViralResultXLS(
            List<LabResultSummary> viralLoaderResultSummary, List<LabResults> viralLoadResults,
            List<LabResults> unsyncronizedViralLoadResults,
            List<PendingHealthFacilitySummary> pendingHealthFacilitySummaries) {

        try (InputStream in = new ClassPathResource("templates/SyncReport.xlsx").getInputStream();
                XSSFWorkbook workbook = new XSSFWorkbook(in);
                ByteArrayOutputStream stream = new ByteArrayOutputStream();) {

            composeVariablesSheet(workbook);
            composeReceivedByDistrictSheet(viralLoaderResultSummary, workbook);
            composeReceivedByUSSheet(viralLoaderResultSummary, workbook);
            composeReceivedByNIDSheet(viralLoadResults, workbook);
            composePendingByUSSheet(pendingHealthFacilitySummaries, workbook);
            composePendingByNIDSheet(unsyncronizedViralLoadResults, workbook);
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(stream);
            return new ByteArrayResource(stream.toByteArray());

        } catch (IOException e) {
            throw new RuntimeException("Could not generate the file", e);
        }
    }

    private void composeVariablesSheet(Workbook workbook) {
        String startDateFormatted = reportDateInterval.getStartDateTime().toLocalDate()
                .format(DATE_FORMAT);
        String endDateFormatted = reportDateInterval.getEndDateTime().toLocalDate()
                .format(DATE_FORMAT);
        Sheet sheet = workbook.getSheetAt(VARIABLES_SHEET);
        int startDateRow = 0;
        int endDateRow = 1;
        sheet.getRow(startDateRow).createCell(1).setCellValue(startDateFormatted);
        sheet.getRow(endDateRow).createCell(1).setCellValue(endDateFormatted);
        workbook.setSheetHidden(VARIABLES_SHEET, true);
    }

    private void composeReceivedByDistrictSheet(List<LabResultSummary> viralLoaderResultSummaryList,
            Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(RECEIVED_BY_DISTRICT_SHEET);

        AtomicInteger counter4 = new AtomicInteger(2);
        Map<String, Map<String, Map<String, ViralResultStatistics>>> provinces = viralLoaderResultSummaryList
                .stream()
                .collect(groupingBy(LabResultSummary::getRequestingProvinceName,
                        groupingBy(LabResultSummary::getRequestingDistrictName,
                                groupingBy(LabResultSummary::getTypeOfResult,
                                        ViralResultStatisticsCollector.toVlResultStatistics()))));

        ViralResultStatistics totals = new ViralResultStatistics();
        provinces.forEach((province, districts) -> {
            districts.forEach((district, typesOfResult) -> {
                typesOfResult.forEach((type, stats) -> {
                    Row row = sheet.createRow(counter4.getAndIncrement());
                    createStatResultRow(workbook, row, province, district, stats);
                    totals.accumulate(stats);
                });
            });
        });

        Row row = sheet.createRow(counter4.getAndIncrement());
        createStatLastResultRow(workbook, row, totals);
    }

    private void composePendingByUSSheet(List<PendingHealthFacilitySummary> pendingViralResultSummaries,
            Workbook workbook) {
        Sheet sheet4 = workbook.getSheetAt(PENDING_BY_US_SHEET);
        AtomicInteger counter = new AtomicInteger(2);
        pendingViralResultSummaries.forEach(pendingViralResultSummary -> {
            Row row = sheet4.createRow(counter.getAndIncrement());
            createPendingViralResultSummaryRow(row, pendingViralResultSummary);
        });
    }

    private void composePendingByNIDSheet(List<LabResults> unsyncronizedViralLoadResults,
            Workbook workbook) {
        Sheet sheet3 = workbook.getSheetAt(PENDING_BY_NID_SHEET);
        int rownum = 2;
        for (LabResults viralResult : unsyncronizedViralLoadResults) {
            Row row = sheet3.createRow(rownum++);
            createUnsyncronizedViralResultRow(row, viralResult);
        }
    }

    private void composeReceivedByNIDSheet(List<LabResults> viralLoadResults, Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(RECEIVED_BY_NID_SHEET);
        int rowNum = 2;
        for (LabResults viralResult : viralLoadResults) {
            createReceivedByNIDRow(sheet.createRow(rowNum++), viralResult);
        }
        for (ResultsReceivedByNid r : ResultsReceivedByNid.values()) {
            sheet.autoSizeColumn(r.ordinal());
        }
    }

    private void composeReceivedByUSSheet(List<LabResultSummary> viralLoaderResultSummary,
            Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(RECEIVED_BY_US_SHEET);
        AtomicInteger counter = new AtomicInteger(3);
        viralLoaderResultSummary.stream().forEach(viralResult -> {
            Row row = sheet.createRow(counter.getAndIncrement());
            createViralResultSummaryRow(row, viralResult);
        });

    }

    private void createViralResultSummaryRow(Row row, LabResultSummary viralLoaderResult) {
        for (ResultsByHFSummary byHfSummary : ResultsByHFSummary.values()) {
            Cell cell = row.createCell(byHfSummary.ordinal());
            switch (byHfSummary) {
                case PROVINCE:
                    cell.setCellValue(viralLoaderResult.getRequestingProvinceName());
                    break;
                case DISTRICT:
                    cell.setCellValue(viralLoaderResult.getRequestingDistrictName());
                    break;
                case HEALTH_FACILITY_CODE:
                    cell.setCellValue(StringUtils.center(viralLoaderResult.getHealthFacilityLabCode(), 11, " "));
                    break;
                case HEALTH_FACILITY_NAME:
                    cell.setCellValue(viralLoaderResult.getFacilityName());
                    break;
                case TYPE_OF_RESULT:
                    cell.setCellValue(viralLoaderResult.getTypeOfResult());
                    break;
                case TOTAL_RECEIVED:
                    cell.setCellValue(viralLoaderResult.getTotalReceived());
                    break;
                case TOTAL_PROCESSED:
                    cell.setCellValue(viralLoaderResult.getProcessed());
                    break;
                case TOTAL_PENDING:
                    cell.setCellValue(viralLoaderResult.getTotalPending());
                    break;
                case NOT_PROCESSED_INVALID_RESULT:
                    cell.setCellValue(viralLoaderResult.getNotProcessedInvalidResult());
                    break;
                case NOT_PROCESSED_NID_NOT_FOUND:
                    cell.setCellValue(viralLoaderResult.getNotProcessedNidNotFount());
                    break;
                case NOT_PROCESSED_DUPLICATED_NID:
                    cell.setCellValue(viralLoaderResult.getNotProcessedDuplicateNid());
                    break;
                case NOT_PROCESSED_DUPLICATED_REQUEST_ID:
                    cell.setCellValue(viralLoaderResult.getNotProcessedDuplicatedRequestId());
                    break;
                default:
                    break;
            }
        }
    }

    private void createReceivedByNIDRow(Row row, LabResults viralLoaderResult) {

        for (ResultsReceivedByNid byNID : ResultsReceivedByNid.values()) {
            Cell cell = row.createCell(byNID.ordinal());
            switch (byNID) {
                case REQUEST_ID:
                    cell.setCellValue(viralLoaderResult.getRequestId());
                    break;
                case TYPE_OF_RESULT:
                    cell.setCellValue(viralLoaderResult.getTypeOfResult());
                    break;
                case NID:
                    cell.setCellValue(viralLoaderResult.getNID());
                    break;
                case PROVINCE:
                    cell.setCellValue(viralLoaderResult.getRequestingProvinceName());
                    break;
                case DISTRICT:
                    cell.setCellValue(viralLoaderResult.getRequestingDistrictName());
                    break;
                case HEALTH_FACILITY_CODE:
                    cell.setCellValue(viralLoaderResult.getHealthFacilityLabCode());
                    break;
                case HEALTH_FACILITY_NAME:
                    cell.setCellValue(viralLoaderResult.getRequestingFacilityName());
                    break;
                case CREATED_AT:
                    cell.setCellValue(viralLoaderResult.getCreatedAt().format(DATE_FORMAT));

                    break;
                case UPDATED_AT:
                    cell.setCellValue(viralLoaderResult.getUpdatedAt() != null
                            ? viralLoaderResult.getUpdatedAt().format(DATE_FORMAT)
                            : "");
                    break;
                case VIRAL_RESULT_STATUS:
                    cell.setCellValue(
                            messageSource.getMessage("disa.viraLoadStatus." + viralLoaderResult.getViralLoadStatus(),
                                    new String[] {}, Locale.getDefault()));
                    break;
                case NOT_PROCESSING_CAUSE:
                    String cellValue = "";
                    if (viralLoaderResult.getNotProcessingCause() != null) {
                        cellValue = messageSource.getMessage(
                                "disa.notProcessingCause." + viralLoaderResult.getNotProcessingCause(), new String[] {},
                                Locale.getDefault());
                    }
                    cell.setCellValue(cellValue);
                    break;
                case OBS:
                    cell.setCellValue(viralLoaderResult.getNotProcessingCause() != null
                            && viralLoaderResult.getNotProcessingCause().trim().equals("NID_NOT_FOUND")
                            && viralLoaderResult.getViralLoadStatus().equals(ViralLoadStatus.PROCESSED.name())
                                    ? "Reprocessado apos a correcao do NID"
                                    : " ");
                    break;
                default:
                    break;
            }

        }
    }

    private void createPendingViralResultSummaryRow(Row row,
            PendingHealthFacilitySummary pendingViralResultSummary) {

        for (ResultsPendingByUs pending : ResultsPendingByUs.values()) {
            Cell cell = row.createCell(pending.ordinal());
            switch (pending) {
                case PROVINCE:
                    cell.setCellValue(pendingViralResultSummary.getRequestingProvinceName());
                    break;
                case DISTRICT:
                    cell.setCellValue(pendingViralResultSummary.getRequestingDistrictName());
                    break;
                case US_CODE:
                    cell.setCellValue(pendingViralResultSummary.getHealthFacilityLabCode());
                    break;
                case US_NAME:
                    cell.setCellValue(pendingViralResultSummary.getFacilityName());
                    break;
                case TOTAL_PENDING:
                    cell.setCellValue(pendingViralResultSummary.getTotalPending());
                    break;
                case LAST_SYNC:
                    cell.setCellValue(pendingViralResultSummary.getLastSyncDate() != null ? pendingViralResultSummary
                            .getLastSyncDate().toLocalDate().format(DATE_FORMAT) : "");
                    break;
                default:
                    break;
            }
        }

    }

    private void createUnsyncronizedViralResultRow(Row row, LabResults viralLoaderResult) {

        for (ResultsPendingByNid pendingByNid : ResultsPendingByNid.values()) {
            Cell cell = row.createCell(pendingByNid.ordinal());
            switch (pendingByNid) {
                case REQUEST_ID:
                    cell.setCellValue(viralLoaderResult.getRequestId());
                    break;
                case NID:
                    cell.setCellValue(viralLoaderResult.getNID());
                    break;
                case PROVINCE:
                    cell.setCellValue(viralLoaderResult.getRequestingProvinceName());
                    break;
                case DISTRICT:
                    cell.setCellValue(viralLoaderResult.getRequestingDistrictName());
                    break;
                case HEALTH_FACILITY_CODE:
                    cell.setCellValue(viralLoaderResult.getHealthFacilityLabCode());
                    break;
                case HEALTH_FACILITY_NAME:
                    cell.setCellValue(viralLoaderResult.getRequestingFacilityName());
                    break;
                case SENT_DATE:
                    cell.setCellValue(
                            viralLoaderResult.getCreatedAt().toLocalDate().format(DATE_FORMAT));
                    break;
                case STATUS:
                    cell.setCellValue(messageSource
                            .getMessage("disa.viraLoadStatus." + viralLoaderResult.getViralLoadStatus(),
                                    new String[] {}, Locale.getDefault()));
                    break;
                default:
                    break;
            }
        }

    }

    private void createStatResultRow(Workbook workbook, Row row, String province, String district,
            ViralResultStatistics viralResultStatistics) {

        for (ResultsByDistrictSummary r : ResultsByDistrictSummary.values()) {
            Cell cell = row.createCell(r.ordinal());
            switch (r) {
                case PROVINCE:
                    cell.setCellValue(province);
                    break;
                case DISTRICT:
                    cell.setCellValue(district);
                    break;
                case TYPE_OF_RESULT:
                    cell.setCellValue(viralResultStatistics.getTypeOfResult());
                    break;
                case TOTAL_PROCESSED:
                    cell.setCellValue(viralResultStatistics.getProcessed());
                    break;
                case PERCENTAGE_PROCESSED:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getProcessedPercentage());
                    break;
                case TOTAL_PENDING:
                    cell.setCellValue(viralResultStatistics.getPending());
                    break;
                case PERCENTAGE_PENDING:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getPendingPercentage());
                    break;
                case NOT_PROCESSED_INVALID_RESULT:
                    cell.setCellValue(viralResultStatistics.getNoProcessedInvalidResult());
                    break;
                case PERCENTAGE_NOT_PROCESSED_INVALID_RESULT:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getNoProcessedNoResultPercentage());
                    break;
                case NOT_PROCESSED_NID_NOT_FOUND:
                    cell.setCellValue(viralResultStatistics.getNoProcessedNidNotFound());
                    break;
                case PERCENTAGE_NOT_PROCESSED_NID_NOT_FOUND:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getNoProcessedNidNotFoundPercentage());
                    break;
                case NOT_PROCESSED_DUPLICATED_NID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicateNid());
                    break;
                case PERCENTAGE_NOT_PROCESSED_DUPLICATED_NID:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicateNidPercentage());
                    break;
                case NOT_PROCESSED_DUPLICATED_REQUEST_ID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicatedReqId());
                    break;
                case PERCENTAGE_NOT_PROCESSED_DUPLICATED_REQUEST_ID:
                    cell.setCellStyle(getPercentCellStyle(workbook));
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicatedReqIdPercentage());
                    break;
                case TOTAL_RECEIVED:
                    cell.setCellValue(viralResultStatistics.getTotal());
                    break;
                default:
                    break;
            }
        }
    }

    private void createStatLastResultRow(Workbook workbook, Row row, ViralResultStatistics viralResultStatistics) {

        for (ResultsByDistrictSummary r : ResultsByDistrictSummary.values()) {
            Cell cell = row.createCell(r.ordinal());
            switch (r) {
                case PROVINCE:
                    cell.setCellValue("Total");
                    cell.setCellStyle(getTotalsCellStyle(workbook));
                    break;
                case TOTAL_PROCESSED:
                    cell.setCellValue(viralResultStatistics.getProcessed());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_PROCESSED:
                    cell.setCellValue(viralResultStatistics.getProcessedPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case TOTAL_PENDING:
                    cell.setCellValue(viralResultStatistics.getPending());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_PENDING:
                    cell.setCellValue(viralResultStatistics.getPendingPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case NOT_PROCESSED_INVALID_RESULT:
                    cell.setCellValue(viralResultStatistics.getNoProcessedInvalidResult());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_NOT_PROCESSED_INVALID_RESULT:
                    cell.setCellValue(viralResultStatistics.getNoProcessedNoResultPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case NOT_PROCESSED_NID_NOT_FOUND:
                    cell.setCellValue(viralResultStatistics.getNoProcessedNidNotFound());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_NOT_PROCESSED_NID_NOT_FOUND:
                    cell.setCellValue(viralResultStatistics.getNoProcessedNidNotFoundPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case NOT_PROCESSED_DUPLICATED_NID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicateNid());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_NOT_PROCESSED_DUPLICATED_NID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicateNidPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case NOT_PROCESSED_DUPLICATED_REQUEST_ID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicatedReqId());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                case PERCENTAGE_NOT_PROCESSED_DUPLICATED_REQUEST_ID:
                    cell.setCellValue(viralResultStatistics.getNotProcessedDuplicatedReqIdPercentage());
                    cell.setCellStyle(getBoldPercentCellStyle(workbook));
                    break;
                case TOTAL_RECEIVED:
                    cell.setCellValue(viralResultStatistics.getTotal());
                    cell.setCellStyle(getBoldStyle(workbook));
                    break;
                default:
                    break;
            }
        }
    }

    private CellStyle getTotalsCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private CellStyle getBoldStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }

    private CellStyle getPercentCellStyle(Workbook workbook) {
        CellStyle percent = workbook.createCellStyle();
        DataFormat df = workbook.createDataFormat();
        percent.setDataFormat(df.getFormat("0%"));
        return percent;
    }

    private CellStyle getBoldPercentCellStyle(Workbook workbook) {
        CellStyle boldPercent = workbook.createCellStyle();
        DataFormat df = workbook.createDataFormat();
        boldPercent.cloneStyleFrom(getBoldStyle(workbook));
        boldPercent.setDataFormat(df.getFormat("0%"));
        return boldPercent;
    }
}
