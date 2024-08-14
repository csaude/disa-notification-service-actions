package disa.notification.service.service;

public class SeafileServiceException extends Exception {

	private static final long serialVersionUID = 1L; 

	public SeafileServiceException(String message) {
		super(message); 
	}
	
	public SeafileServiceException(String message, Throwable cause) {
        super(message, cause);
    }
}
