package disa.notification.service.utils;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.temporal.TemporalField;
import java.time.temporal.WeekFields;
import java.util.Locale;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;

@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class DateTimeUtils {

    private static final int HOUR_23 = 23;
    private static final int SECOND_59 = 59;
    private static final int MINUTE_59 = 59;

    public static DateInterval getLastWeekInterVal() {
        LocalDate now = LocalDate.now();
        // Get the First day of the week
        TemporalField dayOfWeek = WeekFields.of(Locale.US).dayOfWeek();
        LocalDate sunday = now.with(dayOfWeek, 1);
        LocalDate saturdayLastWeek = sunday.minusDays(1);
        LocalDate sundayLastWeek = saturdayLastWeek.minusDays(6);
        return DateInterval.of(sundayLastWeek.atStartOfDay(), saturdayLastWeek.atTime(HOUR_23, MINUTE_59, SECOND_59));
    }

    public static DateInterval getCurrentWeekInterVal() {
        LocalDateTime now = LocalDateTime.now();
        TemporalField dayOfWeek = WeekFields.of(Locale.US).dayOfWeek();
        LocalDate sunday = now.toLocalDate().with(dayOfWeek, 1);
        return DateInterval.of(sunday.atStartOfDay(), now);
    }
}
