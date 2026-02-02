package entities

import kndiyLibraries.DataStructure

import javax.xml.crypto.Data

class Matrix {
    int rowNum
    int columnNum
    int[][] data
    int maxMemberSize
    int maxTabs
    Map valuesByRow
    Map valuesByColumn

    Matrix(int rowNum, int columnNum) {
        initiateMatrix(rowNum, columnNum)
    }

    Matrix(int rowNum, int columnNum, List<List<Integer>> data) {
        initiateMatrix(rowNum, columnNum)
        validateInputs(data, rowNum, columnNum)
        int curRow
        for (int i = 0; i < rowNum; i ++) {
            curRow = i
            List valuesInRow = DataStructure.getOrCreateObject(valuesByRow, i as String, DataStructure.DATA_STRUCTURE_LIST)
            for (int j = 0; j < columnNum; j ++) {
                List valuesInCol = DataStructure.getOrCreateObject(valuesByColumn, j as String, DataStructure.DATA_STRUCTURE_LIST)
                int value = data[i][j] ?: 0
                setValue(i, j, value)
                setMaxMemberSize(value)
                if (curRow == i) {
                    valuesInRow << value
                }
                valuesInCol << value
            }
            curRow ++
        }
    }

    private void setMaxMemberSize(int value) {
        int memberSize = value.toString().size()
        if (memberSize == maxMemberSize) {
            return
        }
        maxMemberSize = [memberSize, maxMemberSize].max()
        maxTabs = ((maxMemberSize / 4) as int) + 1
    }

    private void initiateMatrix(int rowNum, int columnNum) {
        this.rowNum = rowNum
        this.columnNum = columnNum
        this.data = new int[rowNum][columnNum]
        this.valuesByColumn = [ : ]
        this.valuesByRow = [ : ]
    }

    private void validateInputs(List<List<Integer>>data, int rowNum, int columnNum) {
        if (data.size() != rowNum) {
            throw new Exception("Wrong Row number ${rowNum} vs Actual ${data.size()}")
        }
        for (int i = 0; i < rowNum; i++) {
            int[] row = data[i]
            if (row.size() != columnNum) {
                throw new Exception("Wrong Column Size at Row ${i + 1}")
            }
        }
    }

    int getValue(int row, int column) {
        return data[row][column]
    }

    Matrix setValue(int row, int column, int value) {
        if (!data[row][column]) {
            List valuesInRow = DataStructure.getOrCreateObject(valuesByRow, row as String, DataStructure.DATA_STRUCTURE_LIST)
            List valuesInCol = DataStructure.getOrCreateObject(valuesByColumn, column as String, DataStructure.DATA_STRUCTURE_LIST)
            valuesInRow << value
            valuesInCol << value
        }

        data[row][column] = value
        setMaxMemberSize(value)

        return this
    }

    Matrix transposeMatrix() {
        int transposedRowNum = columnNum
        int transposedColumnNum = rowNum
        Matrix transposedMatrix = new Matrix(transposedRowNum, transposedColumnNum)

        for (int i = 0; i < transposedRowNum; i ++) {
            for (int j = 0; j < transposedColumnNum; j ++) {
                transposedMatrix.setValue(i, j, getValue(j, i))
            }
        }

        return transposedMatrix
    }

    Matrix multiply(Matrix multiplierMatrix) {
        if (!verifyMultiplierAndMultiplicand(this, multiplierMatrix)) {
            throw new Exception("Incompatable Matrix in Multiplication!")
        }
        Matrix product = new Matrix(this.rowNum, multiplierMatrix.getColumnNum())
        for (int i = 0 ; i < product.getRowNum(); i ++) {
            for (int j = 0 ; j < product.getColumnNum(); j ++) {
                List candValuesInRow = this.getValuesByRow()?.getAt(i as String)
                List plierValuesInCol = multiplierMatrix?.getValuesByColumn()?.getAt(j as String)
                println(candValuesInRow)
                println(plierValuesInCol)
                int value = 0
                for (int z = 0; z < candValuesInRow.size(); z ++) {
                    value += (candValuesInRow[z] * plierValuesInCol[z])
                }
                product.setValue(i, j, value)
            }
        }

        return product
    }

    boolean verifyMultiplierAndMultiplicand(Matrix multiplicandMatrix, Matrix multiplierMatrix) {
        int candRows = multiplicandMatrix.getRowNum()
        int plierColumns = multiplierMatrix.getColumnNum()
        if (candRows != plierColumns) {
            return false
        }

        int candColumns = multiplicandMatrix.getColumnNum()
        int plierRows = multiplierMatrix.getRowNum()
        if (candColumns != plierRows) {
            return false
        }

        return true
    }

    void printMatrix(String initText = null) {
        String string = initText ? "${initText} = " : ""
        string += "[\n\t"
        for (int i = 0; i < rowNum; i ++) {
            String row = ""
            for (int j = 0; j < columnNum; j ++) {
                int value = getValue(i, j)
                row += "${value}${getTabsString(value)}"
            }
            string += "${row}\n"
            if (i != rowNum - 1) {
                string += "\t"
            }
        }
        string += "]"
        println(string)
    }

    private String getTabsString(int value) {
        int sizeOfCurrentValue = value.toString().size()
        int curTabs = ((sizeOfCurrentValue / 4) as int) + 1

        int requiredTabs = maxTabs - curTabs + 1
        String tabs = ""
        while (requiredTabs > 0) {
            tabs += "\t"
            requiredTabs --
        }

        return tabs
    }
}
