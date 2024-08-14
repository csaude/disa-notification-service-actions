package disa.notification.service.config;

import java.util.Collections;

import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.context.MessageSource;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.client.RestTemplate;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.spring5.SpringTemplateEngine;
import org.thymeleaf.templatemode.TemplateMode;
import org.thymeleaf.templateresolver.ClassLoaderTemplateResolver;
import org.thymeleaf.templateresolver.ITemplateResolver;

import disa.notification.service.service.SeafileService;
import disa.notification.service.service.impl.FileSystemMailService;
import disa.notification.service.service.impl.MailServiceImpl;
import disa.notification.service.service.interfaces.MailService;
import disa.notification.service.utils.DateInterval;
import disa.notification.service.utils.DateTimeUtils;

@Configuration
public class MailSenderConfig {

	@Bean
	public TemplateEngine emailTemplateEngine() {
		final SpringTemplateEngine templateEngine = new SpringTemplateEngine();
		templateEngine.addTemplateResolver(htmlTemplateResolver());
		return templateEngine;
	}

	private ITemplateResolver htmlTemplateResolver() {
		final ClassLoaderTemplateResolver templateResolver = new ClassLoaderTemplateResolver();
		templateResolver.setOrder(Integer.valueOf(2));
		templateResolver.setResolvablePatterns(Collections.singleton("html/*"));
		templateResolver.setPrefix("/mail/");
		templateResolver.setSuffix(".html");
		templateResolver.setTemplateMode(TemplateMode.HTML);
		templateResolver.setCharacterEncoding("spring.mail.encoding");
		templateResolver.setCacheable(false);
		return templateResolver;
	}

	@Bean
	@ConditionalOnProperty(name = "app.mailservice", havingValue = "javaMail")
	public MailService mailServiceImpl(TemplateEngine templateEngine, MessageSource messageSource,
			DateInterval reportDateInterval, SeafileService seafileService) {
		return new MailServiceImpl(templateEngine, messageSource, reportDateInterval, seafileService);
	}

	@Bean
	@ConditionalOnProperty(name = "app.mailservice", havingValue = "fileSystem")
	MailService fileSystemMailService(MessageSource messageSource, DateInterval reportDateInterval) {
		return new FileSystemMailService(messageSource, reportDateInterval);
	}

	@Bean
	@ConditionalOnProperty(name = "app.reportDateInterval", havingValue = "lastWeek")
	DateInterval lastWeekDateInterval() {
		return DateTimeUtils.getLastWeekInterVal();
	}

	@Bean
	@ConditionalOnProperty(name = "app.reportDateInterval", havingValue = "currentWeek")
	DateInterval currentWeekDateInterval() {
		return DateTimeUtils.getCurrentWeekInterVal();
	}
	
	@Bean
	public RestTemplate restTemplate() {
		return new RestTemplate();
	}
}
