import org.joda.time.DateTime;
import org.joda.time.LocalDate;
import org.junit.Test;

import java.util.Calendar;

public class JodaTimeTest {
    @Test
    public void t1() {
        LocalDate now = new LocalDate();
        LocalDate lastDayOfPreviousMonth = now.minusMonths(1).dayOfMonth().withMaximumValue();
        System.out.println(lastDayOfPreviousMonth.toString());
    }

    @Test
    public void t2() {
        // Use a Calendar
        java.util.Calendar calendar = Calendar.getInstance();
        DateTime dateTime = new DateTime(calendar);
        System.out.println(dateTime);
        System.out.println(dateTime.toString("yyyy-MM-dd hh:mm:ss"));
        // Use another Joda DateTime
        String timeString = "2006-01-26";
        dateTime = new DateTime(timeString);
        System.out.println("dateTime = " + dateTime);
        System.out.println("dateTime.toString(\"yyyy-MM-dd\") = " + dateTime.toString("yyyy-MM-dd"));
    }
}
