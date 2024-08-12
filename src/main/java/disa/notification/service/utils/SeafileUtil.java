package disa.notification.service.utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

public class SeafileUtil {

	public static String extractJson(String response) {
		
        int startIndex = response.indexOf("{");
        int endIndex = response.lastIndexOf("}");
        
        if (startIndex != -1 && endIndex != -1) {
            return response.substring(startIndex, endIndex + 1);
        } else {
            return null; // JSON not found
        }
    }
	
	public static JsonNode parseJson(String response) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        return objectMapper.readTree(response);
    }
	
	public static String getUploadLink(String response) {
		int startIndex = response.indexOf("https://");
		int endIndex = response.lastIndexOf("\"");
		return response.substring(startIndex, endIndex);
	}
}
