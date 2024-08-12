package disa.notification.service.service.impl;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;

import javax.mail.MessagingException;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.MessageSource;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;

import disa.notification.service.entity.ImplementingPartner;
import disa.notification.service.service.SeafileService;
import disa.notification.service.service.interfaces.LabResultSummary;
import disa.notification.service.service.interfaces.LabResults;
import disa.notification.service.service.interfaces.MailService;
import disa.notification.service.service.interfaces.PendingHealthFacilitySummary;
import disa.notification.service.utils.DateInterval;
import disa.notification.service.utils.ExcelUtil;
import disa.notification.service.utils.MultipartUtil;
import disa.notification.service.utils.SyncReport;
import disa.notification.service.utils.TemplateEngineUtils;
import lombok.extern.log4j.Log4j2;

@Log4j2
public class MailServiceImpl implements MailService {

    public static final String EMAIL_SUBJECT = "Relatório de Sincronização de resultados lab de %s a %s";

    private TemplateEngine templateEngine;
    private final MessageSource messageSource;
    private DateInterval reportDateInterval;
    private final String startDateFormatted;
    private final String endDateFormatted;
    private final SeafileService seafileService;
    
    public MailServiceImpl(TemplateEngine templateEngine, MessageSource messageSource, DateInterval reportDateInterval, SeafileService seafileService) {
        this.templateEngine = templateEngine;
        this.messageSource = messageSource;
        this.reportDateInterval = reportDateInterval;
        this.seafileService = seafileService;
        
        this.startDateFormatted = this.reportDateInterval.getStartDateTime().toLocalDate().format(DateTimeFormatter.ofPattern("dd-MM-yyyy"));
        this.endDateFormatted = this.reportDateInterval.getEndDateTime().toLocalDate().format(DateTimeFormatter.ofPattern("dd-MM-yyyy"));
    }

    @Value("${spring.mail.username}")
    private String fromEmail;

    @Value("${disa.notifier.rest.endpoint}")
    private String disaNotifierEndPoint;

    public void sendEmail(final ImplementingPartner ip,
            								final List<LabResultSummary> viralLoaders, List<LabResults> viralLoadResults,
            								List<LabResults> unsyncronizedViralLoadResults,
            								List<PendingHealthFacilitySummary> pendingHealthFacilitySummaries) {	
    	
    	Context ctx = prepareEmailContext(viralLoaders);
    	String htmlContent = generateHtmlContent(ctx);
    	String attachmentName = generateAttachmentName(ip);

        try {
        	ByteArrayResource attachment = generateAttachment(viralLoaders, viralLoadResults, unsyncronizedViralLoadResults,
                    pendingHealthFacilitySummaries);
            
            if (attachment != null) {
            	processAttachmentAndSendEmail(ip, attachment, htmlContent, attachmentName);
			}
            
        } catch (IOException e) {
            log.error(e);
        }
    }

    public void sendNoResultsEmail(ImplementingPartner ip)
            throws MessagingException, UnsupportedEncodingException {
        
        Context ctx = new Context(new Locale("pt", "BR"));
        ctx.setVariable("fromDate", startDateFormatted);
        ctx.setVariable("toDate", endDateFormatted);

        String[] mailList = ip.getMailListItems();
        templateEngine = TemplateEngineUtils.getTemplateEngine();
        final String htmlContent = this.templateEngine.process("noResults.html", ctx);
        sendEmailHelper(mailList, htmlContent, "notification", null, startDateFormatted, endDateFormatted, ip.getRepoLink(),Boolean.FALSE); 
    }

    private void sendEmailHelper(String[] mailList, String htmlContent, String module,
            String attachmentName, String startDateFormatted, String endDateFormatted, String repoLink, Boolean resultFlag) {

        String subject = String.format(EMAIL_SUBJECT, startDateFormatted, endDateFormatted);
        ResponseEntity<String> emailResult = null;

        try {

            emailResult = MultipartUtil.sendMultipartRequest(disaNotifierEndPoint, mailList,
                    subject, htmlContent, module, attachmentName, startDateFormatted, endDateFormatted, repoLink, resultFlag);

            if (emailResult != null && emailResult.getStatusCode().is2xxSuccessful()) {
                log.info("Email sent successfully");
            } else if (emailResult != null) {
                log.error("Failed to send email. Response code: " + emailResult.getStatusCode());
            }

        } catch (IOException e) {
            log.error(e);
        }

    }
    
    private Context prepareEmailContext(List<LabResultSummary> viralLoaders) {
        final Context ctx = new Context(new Locale("pt", "BR"));
        ctx.setVariable("fromDate", startDateFormatted);
        ctx.setVariable("toDate", endDateFormatted);
        ctx.setVariable("viralLoaders", viralLoaders);
        return ctx;
	}
    
	private String generateHtmlContent(Context ctx) {
		templateEngine = TemplateEngineUtils.getTemplateEngine();
		return templateEngine.process("index.html", ctx);
	}
	
	private String generateAttachmentName(ImplementingPartner ip) {
        return "CSaude_Resultados_SI-SESP_" + ip.getOrgName().toUpperCase() + "_" + startDateFormatted + "_a_" + endDateFormatted + ".xlsx";
	}
	
	private ByteArrayResource generateAttachment(List<LabResultSummary> viralLoaders, List<LabResults> viralLoadResults,
			List<LabResults> unsyncronizedViralLoadResults,
			List<PendingHealthFacilitySummary> pendingHealthFacilitySummaries) throws IOException {  
			SyncReport syncReport = new SyncReport(messageSource, reportDateInterval);
			return syncReport.getViralResultXLS(viralLoaders, viralLoadResults, 
                                            				unsyncronizedViralLoadResults, 
                                            				pendingHealthFacilitySummaries);
	}
	
	private void processAttachmentAndSendEmail(ImplementingPartner ip, ByteArrayResource attachment, String htmlContent, String attachmentName) {
		String[] mailList = ip.getMailList().split(",");
		try {
			ExcelUtil.saveWorkbook(attachment, attachmentName); 
			seafileService.uploadFile(ip.getRepoId(),attachmentName);
            sendEmailHelper(mailList, htmlContent, "notification",
            				attachmentName,startDateFormatted,endDateFormatted,ip.getRepoLink(),Boolean.TRUE);
            deleteTemporaryFile(attachmentName);
		} catch (Exception e) {
            log.error("Error processing attachment and sending email", e);
        }
	}

	private void deleteTemporaryFile(String attachmentName) throws IOException {
		Files.deleteIfExists(Paths.get("temp").resolve(Paths.get(attachmentName)));
	}
}
