package services.readTxtNovelParsingService


import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun

class DocxPrinter {
    NovelConverter NOVEL_CONVERTER
    XWPFDocument OUTPUT_DOC
    XWPFParagraph CURRENT_PARAGRAPH
    int TOTAL_CHAPTER_ADDED

    DocxPrinter(NovelConverter novelConverter) {
        NOVEL_CONVERTER = novelConverter
        TOTAL_CHAPTER_ADDED = 0

        OUTPUT_DOC = createOutputDocFromTemplate()
        fillOutputDoc()
        printOutputDoc()
    }


    void printOutputDoc() {
        try (FileOutputStream outputStream = new FileOutputStream(NOVEL_CONVERTER.getTARGET_PATH())) {
            OUTPUT_DOC.write(outputStream)
        }
        catch (Exception exception) {
            println(exception.getMessage())
        }
        finally {
            OUTPUT_DOC.close()

            List faultyChapters = findFaultyChapters()
            println("${faultyChapters.size()} Fautly Chapters: ${faultyChapters}")
        }
    }

    private XWPFDocument createOutputDocFromTemplate() {
        XWPFDocument document = new XWPFDocument(new FileInputStream(NOVEL_CONVERTER.getTEMPLATE_PATH()))
        document.removeBodyElement(0)

        return document
    }

    private List findFaultyChapters() {
        List faultyChapters = [ ]
        Integer intChapNum = 0
        NOVEL_CONVERTER.CONTENT_BY_CHAP_NUMBER.each { String chapterNumber, Map content ->
            Integer currIntChapNum = chapterNumber.toBigDecimal().toInteger()
            if (!content?.getAt(NOVEL_CONVERTER.PARAGRAPHS)) {
                faultyChapters.add("EmptyBody:${currIntChapNum}")
            }
            if (currIntChapNum - intChapNum > 1) {
                faultyChapters.add("SkippedChapter:${intChapNum}")
            }
            intChapNum = currIntChapNum
        }

        return faultyChapters
    }

    private void fillOutputDoc() {
        NOVEL_CONVERTER.CONTENT_BY_CHAP_NUMBER.each { String chapterNumber, Map content ->
            String chapterName = content?.getAt(NOVEL_CONVERTER.CHAPTER_PARAGRAPH)
            println("WRITING CHAPTER ${chapterNumber} :: ${chapterName}")
            addChapterParagraph(chapterName)
            addAuthorParagraph()

            List body = content?.getAt(NOVEL_CONVERTER.PARAGRAPHS)
            addBodyParagraph(body)
        }
    }

    private void addChapterParagraph(String chapterName) {
        CURRENT_PARAGRAPH = createParagraph(chapterName, 14, true)
        CURRENT_PARAGRAPH.setStyle("Heading1")

        if (TOTAL_CHAPTER_ADDED != 0) {
            CURRENT_PARAGRAPH.setPageBreak(true)
        }
        TOTAL_CHAPTER_ADDED ++
    }

    private void addAuthorParagraph() {
        String authorParagraph = "${NOVEL_CONVERTER.AUTHOR_NAME} - ${NOVEL_CONVERTER.BOOK_NAME}"
        CURRENT_PARAGRAPH = createParagraph(authorParagraph, 14, true, true, UnderlinePatterns.SINGLE)
    }

    void addBodyParagraph(List body) {
        for (String bodyPara : body) {
            CURRENT_PARAGRAPH = createParagraph(bodyPara)
            CURRENT_PARAGRAPH.setStyle("Normal")
        }
    }

    private XWPFParagraph createParagraph(String text,
                                          int fontSize = 13,
                                          boolean isBold = false,
                                          boolean isItalic = false,
                                          UnderlinePatterns underlinePatterns = UnderlinePatterns.NONE,
                                          ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT,
                                          String fontFamily = "SimSun") {
        XWPFParagraph paragraph = OUTPUT_DOC.createParagraph()
        paragraph.setAlignment(paragraphAlignment)

        XWPFRun run = paragraph.createRun()
        run.setText(text)
        run.setFontFamily(fontFamily)
        run.setFontSize(fontSize)
        run.setBold(isBold)
        run.setItalic(isItalic)
        run.setUnderline(underlinePatterns)

        return paragraph
    }
}
