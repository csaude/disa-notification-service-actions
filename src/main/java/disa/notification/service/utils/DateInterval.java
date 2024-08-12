package disa.notification.service.utils;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.time.LocalDateTime;

@Data
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public class DateInterval {
    private LocalDateTime startDateTime;
    private  LocalDateTime endDateTime;

    public static  DateInterval of(LocalDateTime startDateTime,LocalDateTime endDateTime){
        return new DateInterval(startDateTime,endDateTime);
    }
}
