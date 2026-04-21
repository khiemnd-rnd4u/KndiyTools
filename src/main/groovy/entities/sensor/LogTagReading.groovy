package entities.sensor

import services.kndiyLibraries.DateTimeResolver

import java.time.ZonedDateTime

class LogTagReading {
    private ZonedDateTime dateTime
    private BigDecimal temperature
    private BigDecimal humidity
    private String event

    LogTagReading(ZonedDateTime dateTime,
                  BigDecimal temperature,
                  BigDecimal humidity,
                  String event = null) {
        this.dateTime = dateTime
        this.temperature = temperature
        this.humidity = humidity
        this.event = event
    }

    ZonedDateTime getDateTime() {
        return dateTime
    }

    BigDecimal getTemperature() {
        return temperature
    }

    BigDecimal getHumidity() {
        return humidity
    }

    String getEvent() {
        return event
    }

    @Override
    String toString() {
        return "LogTagReading{" +
                "dateTime=" + dateTime +
                ", temperature=" + temperature +
                ", humidity=" + humidity +
                ", event='" + event + '\'' +
                '}';
    }
}
