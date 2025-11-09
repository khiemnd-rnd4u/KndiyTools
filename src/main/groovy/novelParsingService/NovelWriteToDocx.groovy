package novelParsingService

import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun

class NovelWriteToDocx {
    XWPFDocument outputDoc
    XWPFParagraph currentParagraph
    NovelProcessor novelProcessor
    int totalChapterParagraphAdded

    NovelWriteToDocx(NovelProcessor novelProcessor) {
        this.novelProcessor = novelProcessor
        totalChapterParagraphAdded = 0
    }

    void printOutputDoc() {
        createOutputDocFromTemplate()
        fillOutputDoc()

        try (FileOutputStream outputStream = new FileOutputStream(novelProcessor.outputPath)) {
            outputDoc.write(outputStream)
        }
        catch (Exception exception) {
            println(exception.getMessage())
        }
        finally {
            outputDoc.close()

            List faultyChapters = findFaultyChapters()
            println("${faultyChapters.size()} Fautly Chapters: ${faultyChapters}")
        }
    }

    private void createOutputDocFromTemplate() {
        outputDoc = new XWPFDocument(new FileInputStream(novelProcessor.templatePath))
        outputDoc.removeBodyElement(0)
    }

    private List findFaultyChapters() {
        List faultyChapters = [ ]
        Integer intChapNum = 0
        novelProcessor.paragraphsByTypeByChapterNumber.each { chapterNumber, paragraphsByType ->
            Integer currIntChapNum = chapterNumber.toBigDecimal().toInteger()
            if (!paragraphsByType?.getAt("body")) {
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
        novelProcessor.paragraphsByTypeByChapterNumber.each { String chapterNumber, paragraphsByType ->
            List title = paragraphsByType?.getAt("title")
            if (!title) {
                return
            }
            String chapterName = title[0]
            println("WRITING CHAPTER ${chapterNumber} :: ${chapterName}")
            addBookNameAndChapterParagraph(chapterName)

            String author = title[1]
            addAuthorParagraph(author)

            List body = paragraphsByType?.getAt("body")
            addBodyParagraph(body)
        }
    }

    private void addBookNameAndChapterParagraph(String chapterName) {
        currentParagraph = createParagraph(chapterName, 14, true)
        currentParagraph.setStyle("Heading1")

        if (totalChapterParagraphAdded != 0) {
            currentParagraph.setPageBreak(true)
        }
        totalChapterParagraphAdded ++
    }

    private void addAuthorParagraph(String author) {
        currentParagraph = createParagraph(author, 14, false, true, UnderlinePatterns.SINGLE)
    }

    void addBodyParagraph(List body) {
        for (String bodyPara : body) {
            currentParagraph = createParagraph(bodyPara)
            currentParagraph.setStyle("Normal")
        }
    }

    private XWPFParagraph createParagraph(String text,
                                  int fontSize = 13,
                                  boolean isBold = false,
                                  boolean isItalic = false,
                                  UnderlinePatterns underlinePatterns = UnderlinePatterns.NONE,
                                  ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT,
                                  String fontFamily = "Bookerly") {
        XWPFParagraph paragraph = outputDoc.createParagraph()
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
