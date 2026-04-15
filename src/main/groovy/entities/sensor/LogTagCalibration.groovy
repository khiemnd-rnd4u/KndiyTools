package entities.sensor

import services.kndiyLibraries.DateTimeResolver

import java.time.ZonedDateTime

class LogTagCalibration {
    static final String CI_CALIBRATION_ID_BY_LT_SERIAL = "CICalibrationIdByLtSerial"
    static final String CI_CALIBRATION_DATE = "CICalibrationDate"
    static final String CI_CALIBRATION_DUE_DATE = "CICalibrationDueDate"
    private static final Map<ZonedDateTime, Map> CALIBRATION_INFO_BY_DATE = [
            (DateTimeResolver.getZonedDateTime("2025-12-25")): [
                    (CI_CALIBRATION_ID_BY_LT_SERIAL) : LogTagCalibration20251225.getCalibrationIdByLtSerial(),
                    (CI_CALIBRATION_DATE) : LogTagCalibration20251225.getCalibrationDate(),
                    (CI_CALIBRATION_DUE_DATE) : LogTagCalibration20251225.getCalibrationDueDate(),
            ]
    ]
    private static final List CALIBRATION_DATES = (CALIBRATION_INFO_BY_DATE?.keySet() as List).sort()

    private String logTagSerial
    private String calibrationId
    private ZonedDateTime calibrationDate
    private ZonedDateTime calibrationDueDate
    private ZonedDateTime inspectionDate

    LogTagCalibration(String logTagSerial,
                      ZonedDateTime inspectionDate) {
        this.inspectionDate = inspectionDate
        this.logTagSerial = logTagSerial
        findMatchedCalibrationDate()
        findDueDate()
        findCalibrationId()
    }

    private void findMatchedCalibrationDate() {
        for (ZonedDateTime calibrationDate : CALIBRATION_DATES) {
            Map calibrationInfo = CALIBRATION_INFO_BY_DATE?.getAt(calibrationDate)
            ZonedDateTime dueDate = calibrationInfo?.getAt(CI_CALIBRATION_DUE_DATE)
            if (inspectionDate >= calibrationDate && inspectionDate <= dueDate) {
                this.calibrationDate = calibrationDate
                return
            }
        }
    }

    private void findDueDate() {
        if (!calibrationDate) {
            return
        }

        Map calibrationInfo = CALIBRATION_INFO_BY_DATE?.getAt(calibrationDate)
        this.calibrationDueDate = calibrationInfo?.getAt(CI_CALIBRATION_DUE_DATE)
    }

    private void findCalibrationId() {
        if (!calibrationDate) {
            return
        }
        Map calibrationInfo = CALIBRATION_INFO_BY_DATE?.getAt(calibrationDate)
        this.calibrationId = calibrationInfo?.getAt(CI_CALIBRATION_ID_BY_LT_SERIAL)
                ?.getAt(logTagSerial)
    }

    @Override
    String toString() {
        return "LogTagCalibration{" +
                "logTagSerial='" + logTagSerial + '\'' +
                ", calibrationId='" + calibrationId + '\'' +
                ", calibrationDate='" + calibrationDate + '\'' +
                ", calibrationDueDate='" + calibrationDueDate + '\'' +
                ", inspectionDate='" + inspectionDate + '\'' +
                '}';
    }

    String getLogTagSerial() {
        return logTagSerial
    }

    String getCalibrationId() {
        return calibrationId
    }

    ZonedDateTime getCalibrationDate() {
        return calibrationDate
    }

    ZonedDateTime getCalibrationDueDate() {
        return calibrationDueDate
    }

    ZonedDateTime getInspectionDate() {
        return inspectionDate
    }
}
