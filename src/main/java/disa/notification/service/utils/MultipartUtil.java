package disa.notification.service.utils;

import java.io.IOException;

import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

import disa.notification.service.model.EmailDTO;
import lombok.extern.log4j.Log4j2;

@Log4j2
public class MultipartUtil {

	public static ResponseEntity<String> sendMultipartRequest(String url, String[] mailList,
			String subject, String body, String module, String attachmentName, String startDate,
			String endDate, String repoLink, Boolean resultFlag) throws IOException {

		RestTemplate restTemplate = new RestTemplate();

		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.MULTIPART_FORM_DATA);

		MultiValueMap<String, Object> requestBody = new LinkedMultiValueMap<>();

		// JSON Part
		HttpHeaders jsonHeaders = new HttpHeaders();
		jsonHeaders.setContentType(MediaType.APPLICATION_JSON);

		EmailDTO emailDto = new EmailDTO();
		emailDto.setTo(mailList);
		emailDto.setSubject(subject);
		emailDto.setBody(body);
		emailDto.setModule(module);
		emailDto.setStartDate(startDate);
		emailDto.setEndDate(endDate);
		emailDto.setEndDate(endDate);
		emailDto.setRepoLink(repoLink);
		emailDto.setResultFlag(resultFlag); 

		HttpEntity<EmailDTO> jsonEntity = new HttpEntity<>(emailDto, jsonHeaders);
		requestBody.add("data", jsonEntity);

		HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(requestBody, headers);

		try {
			return restTemplate.exchange(url, HttpMethod.POST, requestEntity, String.class);
		} catch (RestClientException e) {
			log.error(e);
			return null;
		}
	}
}
