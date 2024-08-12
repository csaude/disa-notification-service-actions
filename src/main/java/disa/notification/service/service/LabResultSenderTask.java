package disa.notification.service.service;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Set;

import javax.mail.MessagingException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import disa.notification.service.entity.ImplementingPartner;
import disa.notification.service.repository.ImplementingPartnerRepository;
import disa.notification.service.repository.ViralLoaderRepository;
import disa.notification.service.service.interfaces.LabResultSummary;
import disa.notification.service.service.interfaces.LabResults;
import disa.notification.service.service.interfaces.MailService;
import disa.notification.service.service.interfaces.PendingHealthFacilitySummary;
import disa.notification.service.utils.DateInterval;
import lombok.RequiredArgsConstructor;

@Component
@RequiredArgsConstructor
public class LabResultSenderTask {
    private static final Logger log = LoggerFactory.getLogger(LabResultSenderTask.class);

    private final ImplementingPartnerRepository ipRepository;
    private final MailService mailService;
    private final DateInterval reportDateInterval;
    private final ViralLoaderRepository viralLoaderRepository;

    @Scheduled(cron = "${task.cron}")
    public void sendLabResultReport() {
        log.info("Iniciando a task de Sincronizacao de Cargas virais");
        log.info("Report date interval {}", reportDateInterval);
        log.info("A Compor Dados para envio");

        // Custom query method that returns all implementing entities where the enabled
        // field is true,
        // and it uses the @EntityGraph annotation to ensure that related entities are
        // loaded along with
        // the query results
        List<ImplementingPartner> implementingPartners = ipRepository.findByEnabledTrue();

        for (ImplementingPartner implementingPartner : implementingPartners) {
            log.info(" A Sincronizar Dados da Provincia de {}", implementingPartner.getOrgName());
            sendEmailForImplementingPartner(implementingPartner);
        }

    }

    private void sendEmailForImplementingPartner(ImplementingPartner implementingPartner) {
        LocalDateTime startDateTime = reportDateInterval.getStartDateTime();
        LocalDateTime endDateTime = reportDateInterval.getEndDateTime();
        Set<String> orgUnitCodes = implementingPartner.getOrgUnitCodes();
        List<LabResultSummary> labResultSummary = viralLoaderRepository.findViralLoadResultSummary(
                startDateTime,
                endDateTime,
                orgUnitCodes);
        List<LabResults> labResults = viralLoaderRepository.findViralLoadResults(
                startDateTime,
                endDateTime,
                orgUnitCodes);
        List<LabResults> pendingResultsForMoreThan2Days = viralLoaderRepository
                .findViralLoadResultsPendingMoreThan2Days(orgUnitCodes);
        List<PendingHealthFacilitySummary> pendingHealthFacilitySummaries = viralLoaderRepository
                .findUnsincronizedHealthFacilities(orgUnitCodes);

        try {
            if (!labResultSummary.isEmpty() || !pendingResultsForMoreThan2Days.isEmpty()) {
                mailService.sendEmail(implementingPartner, labResultSummary, labResults, pendingResultsForMoreThan2Days,
                        pendingHealthFacilitySummaries);
            } else {
                mailService.sendNoResultsEmail(implementingPartner);
            }
        } catch (IOException | MessagingException e) {
            log.error("Erro ao enviar relatório de Cargas virais", e);
        }
    }
}
