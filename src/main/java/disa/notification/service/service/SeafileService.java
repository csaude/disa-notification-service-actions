package disa.notification.service.service;

import java.nio.file.Paths;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

import com.fasterxml.jackson.core.JsonProcessingException;

import disa.notification.service.utils.SeafileUtil;
import lombok.extern.log4j.Log4j2;

@Log4j2
@Service
public class SeafileService {
	
	@Value("${disa.notifier.seafile.username}")
 	private String userName;
	
	@Value("${disa.notifier.seafile.password}") 
	private String password;
	
	@Value("${disa.notifier.seafile.url}") 
	private String seafileUrl;

	public final String HEADER = "--header 'accept: application/json' ";
	
	private final RestTemplate restTemplate;
	
	public SeafileService(RestTemplate restTemplate) {
        this.restTemplate = restTemplate;
    }
	
	public String getAuthenticationToken() throws SeafileServiceException, JsonProcessingException {  
		String url = seafileUrl + "/api2/auth-token/";
		
		HttpHeaders headers = createJsonHeaders();
        
        Map<String, String> requestBody = new HashMap<>();
        requestBody.put("username", userName);
        requestBody.put("password", password);
        
        HttpEntity<Map<String, String>> entity = new HttpEntity<>(requestBody, headers);

        try {
        	ResponseEntity<String> response = restTemplate.exchange(url, HttpMethod.POST, entity, String.class);
        	String jsonResponse = SeafileUtil.extractJson(response.getBody());
        	String token = SeafileUtil.parseJson(jsonResponse).get("token").toString();
        	return token.replace("\"", "");
		} catch (RestClientException e) {
			log.error("Error obtaining authentication token: ", e);
			throw new SeafileServiceException("Failed to get authentication token", e); 
		}
	}
	
	public String getAnUploadLink(String repoId) throws SeafileServiceException, JsonProcessingException {
		String url = UriComponentsBuilder.fromHttpUrl(seafileUrl + "/api2/repos/" + repoId + "/upload-link/")
                .queryParam("p", "/")
                .toUriString();
		
		HttpHeaders headers = createJsonHeaders();
        headers.set("authorization", "Token " + getAuthenticationToken());
        
        HttpEntity<String> entity = new HttpEntity<>(headers);
        
        try {
        	ResponseEntity<String> response = restTemplate.exchange(url, HttpMethod.GET, entity, String.class);
        	return SeafileUtil.getUploadLink(response.getBody());
		} catch (RestClientException e) {
			log.error("Error getting upload link: ", e);
            throw new SeafileServiceException("Failed to get upload link", e);
		}
	}
	
	public String uploadFile(String repoId, String attachmentName) throws SeafileServiceException, JsonProcessingException {
		String uploadUrl = getAnUploadLink(repoId) + "?ret-json=1";
		
		HttpHeaders headers = createJsonHeaders();
        headers.set("authorization", "Token " + getAuthenticationToken());
        headers.setContentType(MediaType.MULTIPART_FORM_DATA);
        
        MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
        body.add("parent_dir", "/");
        body.add("file", new FileSystemResource(Paths.get("temp").resolve(attachmentName).toString()));
        
        HttpEntity<MultiValueMap<String, Object>> entity = new HttpEntity<>(body, headers);
        
        try {
        	ResponseEntity<String> response = restTemplate.exchange(uploadUrl, HttpMethod.POST, entity, String.class);
        	return response.getBody();
		} catch (RestClientException e) {
			log.error("Error uploading file: ", e);
            throw new SeafileServiceException("Failed to upload file", e);
		}
	}
	
	private HttpHeaders createJsonHeaders() {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);
        headers.setAccept(Collections.singletonList(MediaType.APPLICATION_JSON));
        return headers;
    }
}
