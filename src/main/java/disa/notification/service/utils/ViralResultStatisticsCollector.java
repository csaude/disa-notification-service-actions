package disa.notification.service.utils;

import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.BinaryOperator;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Collector;

import disa.notification.service.entity.ViralResultStatistics;
import disa.notification.service.service.interfaces.LabResultSummary;

public class ViralResultStatisticsCollector
        implements Collector<LabResultSummary, ViralResultStatistics, ViralResultStatistics> {

    public static ViralResultStatisticsCollector toVlResultStatistics() {
        return new ViralResultStatisticsCollector();
    }

    @Override
    public BiConsumer<ViralResultStatistics, LabResultSummary> accumulator() {
        return (stats, summary) -> {
            stats.accumulate(summary);
        };
    }

    @Override
    public Set<Characteristics> characteristics() {
        return Set.of(Characteristics.UNORDERED);
    }

    @Override
    public BinaryOperator<ViralResultStatistics> combiner() {
        return (stats, other) -> {
            return stats.combine(other);
        };
    }

    @Override
    public Function<ViralResultStatistics, ViralResultStatistics> finisher() {
        return Function.identity();
    }

    @Override
    public Supplier<ViralResultStatistics> supplier() {
        return ViralResultStatistics::new;
    }
}
