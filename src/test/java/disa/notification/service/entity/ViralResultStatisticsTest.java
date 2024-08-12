package disa.notification.service.entity;


import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;


import org.junit.jupiter.api.Test;

public class ViralResultStatisticsTest {
    @Test
    public void testGetProcessedPercentage() {
        // Create some test summaries
        ViralResultStatistics statistics = ViralResultStatistics.builder()
                .total(150)
                .processed(70)
                .pending(30)
                .noProcessedInvalidResult(15)
                .noProcessedNidNotFound(5)
                .build();

        // Assert that the getProcessedPercentage() method returns the correct value
        assertThat(statistics.getProcessedPercentage()).isEqualTo(70.0 / 150.0,
                within(0.001));
    }
}
