package readTxtNovelParsingService

import chineseTextService.ChineseNumberConverter

class NovelConverter {
    final String CHAPTER_NUMBER_TEMPLATE = "000000000000000000"
    final String CHAPTER_PARAGRAPH = "ChapterParagraphs"
    final String PARAGRAPHS = "Paragraphs"

    final String CHAPTER_STRING_FIRST_CHARACTER = "第"
    final String CHAPTER_STRING_LAST_CHARACTER = "章"

    TreeMap CONTENT_BY_CHAP_NUMBER
    Map CURRENT_CONTENT
    boolean IS_BOOK_NAME_IN_CHAPTER_NAME
    boolean IS_NULL_CHAPTER_FOUND
    String CURRENT_LINE
    String CURRENT_CHAPTER_NUMBER
    String SOURCE_PATH
    String TARGET_PATH
    String TEMPLATE_PATH
    String AUTHOR_NAME
    String BOOK_NAME

    NovelConverter(String sourcePath,
                   String targetPath,
                   String templatePath,
                   String authorName, String bookName,
                   boolean isBookNameInChapterName = false) {
        CONTENT_BY_CHAP_NUMBER = [: ]
        IS_BOOK_NAME_IN_CHAPTER_NAME = isBookNameInChapterName
        IS_NULL_CHAPTER_FOUND = true
        SOURCE_PATH = sourcePath
        TARGET_PATH = targetPath
        TEMPLATE_PATH = templatePath
        AUTHOR_NAME = authorName
        BOOK_NAME= bookName

        readFileAndUpdateHolder()
        printHolderIntoDocxFile()
    }

    void readFileAndUpdateHolder() {
        if (!SOURCE_PATH) {
            println("Please specify the path to Source txt file!!!")
            return
        }

        try (BufferedReader bufferedReader = new BufferedReader(new FileReader(SOURCE_PATH))) {
            readFileAndUpdateHolder(bufferedReader)
        }
        catch (ex) {
            println(ex.getMessage())
        }
    }

    void readFileAndUpdateHolder(BufferedReader bufferedReader) {
        int i = 0
        while ((CURRENT_LINE = bufferedReader.readLine()) != null) {
            if (!CURRENT_LINE) {
                continue
            }
            CURRENT_LINE = CURRENT_LINE.trim()

            updateHolder()
        }
    }

    void updateHolder() {
        if (checkIfIsChapterLine()) {
            CURRENT_CHAPTER_NUMBER = getChapterNumber()
            println("Parsing Chapter: ${CURRENT_CHAPTER_NUMBER}")

            CURRENT_CONTENT = CONTENT_BY_CHAP_NUMBER.get(CURRENT_CHAPTER_NUMBER, [ : ])
            CURRENT_CONTENT[ CHAPTER_PARAGRAPH ] = CURRENT_LINE
            CURRENT_CONTENT[ PARAGRAPHS ] = [ ]

            IS_NULL_CHAPTER_FOUND = false

            return
        }
        else if (IS_NULL_CHAPTER_FOUND) {
            CURRENT_CONTENT = CONTENT_BY_CHAP_NUMBER.get(CHAPTER_NUMBER_TEMPLATE, [ : ])
            CURRENT_CONTENT[ CHAPTER_PARAGRAPH ] = "SOURCE FROM:"
            List paragraphs = CURRENT_CONTENT.get(PARAGRAPHS, [ ])
            paragraphs.add(CURRENT_LINE)

            return
        }

        CURRENT_CONTENT[ PARAGRAPHS ] << CURRENT_LINE
    }

    private boolean checkIfIsChapterLine() {
        if (CURRENT_LINE.startsWith(" ")) {
            return false
        }
        String chapterString = CURRENT_LINE.split(" ")[0]
        if (!chapterString.startsWith(CHAPTER_STRING_FIRST_CHARACTER)) {
            return false
        }
        if (!chapterString.endsWith(CHAPTER_STRING_LAST_CHARACTER)) {
            return false
        }

        return true
    }

    private String getChapterNumber() {
        if (IS_BOOK_NAME_IN_CHAPTER_NAME) {
            return calculateChapterNumberWithBookName()
        }

        return calculateChapterNumber()
    }

    private String calculateChapterNumberWithBookName() {
        return null
    }

    private String calculateChapterNumber() {
        String chapterString = CURRENT_LINE.split(" ")[0]
        chapterString = chapterString.substring(1, chapterString.length() - 1)
        chapterString = ChineseNumberConverter.convertToArabicNumber(chapterString).toString()

        return CHAPTER_NUMBER_TEMPLATE.substring(0, CHAPTER_NUMBER_TEMPLATE.length() - chapterString.length()) +
                chapterString
    }

    private void printHolderIntoDocxFile() {
        new DocxPrinter(this)
    }
}
