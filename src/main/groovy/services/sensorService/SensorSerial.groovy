package services.sensorService


import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

class SensorSerial {
    private static final Map SENSOR_SERIAL_BY_ID = [
            "51" : "1010209722",
            "52" : "1010209607",
            "53" : "1010211200",
            "54" : "1010211214",
            "56" : "1010209701",
            "57" : "1010211181",
            "59" : "1010209606",
            "60" : "1010211221",
            "61" : "1010209712",
            "62" : "1010211201",
            "63" : "1010209681",
            "64" : "1010211198",
            "65" : "1010211186",
            "66" : "1010209605",
            "67" : "1010211218",
            "68" : "1010209676",
            "69" : "1010211183",
            "70" : "1010211199",
            "71" : "1010211210",
            "72" : "1010211207",
            "73" : "1010211209",
            "74" : "1010211179",
            "75" : "1010211180",
            "76" : "1010209708",
            "77" : "1010211182",
            "78" : "1010209710",
            "79" : "1010211196",
    ]

    private String FOLDER_PATH
    private String MOVE_PATH
    private Integer SERIAL_IDX
    private Integer CERTIFICATION_NUMBER_IDX
    private String SEPARATOR
    private List TEXTS_TO_REMOVE_FROM_NAME
    private final List FILES
    private Map SENSOR_ID_BY_SERIAL

    SensorSerial(String folderPath,
                 int serialIdx = 0,
                 int certificationNumberIdx = 1,
                 String separator = "_") {
        File folder = new File(folderPath)
        FOLDER_PATH = folderPath
        resolveMovePath()
        SERIAL_IDX = serialIdx
        CERTIFICATION_NUMBER_IDX = certificationNumberIdx
        SEPARATOR = separator

        FILES = folder.listFiles()
        SENSOR_ID_BY_SERIAL = SENSOR_SERIAL_BY_ID
                .collectEntries { String number, String serial ->
                    return [
                            (serial) : number
                    ]
                }
    }

    private String resolveMovePath() {
        if (FOLDER_PATH.endsWith("\\")) {
            MOVE_PATH = "${FOLDER_PATH}RenamedCertifications"
        }
        else {
            MOVE_PATH = "${FOLDER_PATH}\\RenamedCertifications"
        }
    }

    void renameAndMoveFiles() {
        FILES.each { File file ->
            if (!file.getName().endsWith(".pdf")) {
                return
            }

            Path sourceFile = Paths.get(file.getPath())
            Path moveToFolder = Paths.get(MOVE_PATH)
            Files.createDirectories(moveToFolder)

            String fileName = getNewFileName(file)
            Path moveToFile = moveToFolder.resolve(fileName)
            Files.copy(sourceFile, moveToFile)
            println("copy for ${moveToFile}")
        }
    }

    private String getNewFileName(File pdfCertificate) {
        String fileName = pdfCertificate.getName()
        fileName = fileName.replace(".pdf", "")
        List nameCompos = fileName?.split(SEPARATOR)

        String serial = nameCompos?.getAt(SERIAL_IDX)
        String certificationNumber = nameCompos?.getAt(CERTIFICATION_NUMBER_IDX)

        String logTagNumber = SENSOR_ID_BY_SERIAL?.getAt(serial)

        return "${logTagNumber}_${serial}_${certificationNumber}.pdf"
    }
}
