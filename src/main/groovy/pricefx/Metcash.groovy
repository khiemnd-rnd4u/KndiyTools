package pricefx

import kndiyLibraries.DataStructure
import java.nio.file.*

class Metcash {
    Map getIntegrationFileByStateByDsName(String folderPath) {
        Map integrationFileByStateByDsName = [ : ]
        File folder = new File(folderPath)
        println("Parsing Folder: ${folder.getName()}")
        for (File file : folder.listFiles()) {
            println("Parsing File: ${file.getName()}")

            String fileName = file.getName()
            String nameFirstPart = fileName.split("\\.")?.getAt(0)
            String state = nameFirstPart.substring(1, 4)
            String dsName = nameFirstPart.substring(4)

            Map integrationFileByState = DataStructure.getOrCreateObject(integrationFileByStateByDsName, dsName)
            integrationFileByState[ state ] = file
        }

        return integrationFileByStateByDsName
    }

    void copyFilesToFolder(Map integrationFileByStateByDsName, String pathToMove) {
        integrationFileByStateByDsName.each { String dsName, Map fileByState ->
            fileByState.each { String state, File file ->
                Path sourceFile = Paths.get(file.getPath())
                pathToMove = pathToMove.endsWith("\\/")
                        ? "${pathToMove}${dsName}"
                        : "${pathToMove}/${dsName}"
                Path moveToPath = Paths.get(pathToMove)

                String fileName = sourceFile.getFileName()
                fileName.substring(0, fileName.indexOf("__"))
                Path moveToFile = moveToPath.resolve(fileName)

                try {
                    Files.copy(sourceFile, moveToFile)
                    println("Copy completed for ${sourceFile.getFileName()}")
                }
                catch (IOException ex) {
                    println("Copy failed: ${ex.getMessage()}")
                }
            }
        }
    }
}
