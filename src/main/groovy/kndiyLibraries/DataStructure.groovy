package kndiyLibraries

class DataStructure {
    static final String DATA_STRUCTURE_LIST = "List"
    static final String DATA_STRUCTURE_MAP = "Map"
    static final String DATA_STRUCTURE_TREE_MAP = "TreeMap"

    static Object getOrCreateObject(Map objectByKey, String key, String objectType = DATA_STRUCTURE_MAP) {
        Object object = objectByKey[ key ]
        if (object == null) {
            if (objectType == DATA_STRUCTURE_MAP) {
                object = [ : ]
            }
            else if (objectType == DATA_STRUCTURE_LIST) {
                object = [ ]
            }
            else if (objectType == DATA_STRUCTURE_TREE_MAP) {
                object = [ : ] as TreeMap
            }
            objectByKey[ key ] = object
        }

        return object
    }
}
