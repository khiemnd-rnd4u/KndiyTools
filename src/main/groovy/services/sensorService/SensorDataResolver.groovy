package services.sensorService

import entities.sensor.Inspection
import entities.sensor.LogTagSensor

import java.time.ZonedDateTime
import services.kndiyLibraries.DateTimeResolver

class SensorDataResolver {
    private TreeMap<ZonedDateTime, Map> dataByTypeByDateTimeByFileName
    private Inspection inspection
    private Integer locationIdx
    private Integer altitudeIdx
    private Integer serialIdx
    private ZonedDateTime inspectionDate

    SensorDataResolver(TreeMap dataByTypeByDateTimeByFileName,
                       String inspectionId,
                       String inspectionName,
                       String inspectionAddress,
                       String inspectionDate,
                       BigDecimal minTemperatureTolerance = 15,
                       BigDecimal maxTemperatureTolerance = 25,
                       BigDecimal minHumidityTolerance = 0,
                       BigDecimal maxHumidityTolerance = 75,
                       Integer locationIdx = 0,
                       Integer serialIdx = 2,
                       Integer altitudeIdx = null) {
        this.locationIdx = locationIdx
        this.altitudeIdx = altitudeIdx
        this.serialIdx = serialIdx
        this.inspectionDate = DateTimeResolver.getZonedDateTime(inspectionDate)

        this.dataByTypeByDateTimeByFileName = dataByTypeByDateTimeByFileName
        inspection = new Inspection(
                inspectionId, inspectionName, inspectionAddress,
                minTemperatureTolerance, maxTemperatureTolerance,
                minHumidityTolerance, maxHumidityTolerance
        )
        prepareInspection()
    }

    private void prepareInspection() {
        dataByTypeByDateTimeByFileName.each { String fileName , Map dataByTypeByDateTime ->
            LogTagSensor sensor = getSensor(fileName, inspectionDate)
            dataByTypeByDateTime.each { ZonedDateTime dateTime, Map dataByType ->
                addReading(sensor, dataByType, dateTime)
            }
            sensor.resolveAvgAndMkt()

            if (checkIfEvnSensor(fileName)) {
                inspection.addEnvironmentSensor(sensor)
            }
            else {
                inspection.addLogTagSensor(sensor)
            }
        }
    }

    boolean checkIfEvnSensor(String fileName) {
        return fileName?.startsWith("zENV") || fileName.contains("Ngoai")
    }

    private LogTagSensor getSensor(String fileName, ZonedDateTime inspectionDate) {
        List compos = fileName.split("_")
        String idOrSerial = compos[ serialIdx ]
        String location = compos[ locationIdx ]
        String altitude = altitudeIdx ? compos[ altitudeIdx ] : null

        return new LogTagSensor(idOrSerial, location, altitude, inspectionDate)
    }

    private void addReading(LogTagSensor sensor, Map dataByType, ZonedDateTime dateTime) {
        BigDecimal temperature = dataByType?.getAt(SensorDataParser.TEMPERATURE) as BigDecimal
        BigDecimal humidity = dataByType?.getAt(SensorDataParser.HUMIDITY) as BigDecimal
        String event = dataByType?.getAt(SensorDataParser.EVENT)

        sensor.addReading(dateTime, temperature, humidity, event)
    }

    Inspection getInspection() {
        return inspection
    }
}
