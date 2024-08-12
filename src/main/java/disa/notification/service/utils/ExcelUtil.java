package disa.notification.service.utils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;

import lombok.extern.log4j.Log4j2;

@Log4j2
public class ExcelUtil {
	
	public static void saveWorkbook(ByteArrayResource byteArrayResource, String fileName) throws IOException, EncryptedDocumentException, InvalidFormatException {
        byte[] byteArray = byteArrayResource.getByteArray();
        
        // create temp folder if it doesn't exist
        if (Files.notExists(Paths.get("temp"))) { 
        	Files.createDirectory(Paths.get("temp"));
		}
        
        // Converter o ByteArrayResource em um Workbook
        try (ByteArrayInputStream bis = new ByteArrayInputStream(byteArray);
             Workbook workbook = WorkbookFactory.create(bis)) {
             
            // Salvar o Workbook na raiz do projeto
            try (FileOutputStream fos = new FileOutputStream(Paths.get("temp").resolve(fileName).toString())) { 
                workbook.write(fos);
            }
        }
    }

    public static File getFileFromResource(String filename) throws IOException {
        ClassPathResource resource = new ClassPathResource(filename);
        return resource.getFile().getAbsoluteFile();
    }
}
