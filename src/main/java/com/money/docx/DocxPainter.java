package com.money.docx;

import lombok.Getter;
import org.docx4j.Docx4J;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.model.fields.merge.MailMergerWithNext;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.toc.TocException;
import org.docx4j.toc.TocGenerator;
import org.docx4j.wml.P;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.Tbl;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯画家
 * <a href="https://www.docx4java.org/trac/docx4j">docx4j官网</a>
 * @createTime : 2022-08-20 15:27:07
 */
public class DocxPainter {

    @Getter
    private final WordprocessingMLPackage wpk;

    private TocGenerator tocGenerator;

    private Consumer<Exception> exceptionCapture = e -> {
        throw new RuntimeException(e);
    };

    public DocxPainter() {
        try {
            wpk = WordprocessingMLPackage.createPackage();
        } catch (Exception e) {
            throw new RuntimeException("DocxPainter init failure", e);
        }
    }

    public DocxPainter(File file) {
        try {
            wpk = Docx4J.load(file);
        } catch (Exception e) {
            throw new RuntimeException("DocxPainter init failure", e);
        }
    }

    public DocxPainter processTemplate(Map<DataFieldName, String> ctx) {
        try {
            List<Map<DataFieldName, String>> data = Collections.singletonList(ctx);
            MailMerger.setMERGEFIELDInOutput(MailMerger.OutputField.KEEP_MERGEFIELD);
            MailMergerWithNext.performLabelMerge(wpk, data);
        } catch (Docx4JException e) {
            this.exceptionCatch(e);
        }
        return this;
    }

    public DocxPainter add(P p) {
        wpk.getMainDocumentPart().addObject(p);
        return this;
    }

    public DocxPainter add(Tbl tbl) {
        wpk.getMainDocumentPart().addObject(tbl);
        return this;
    }

    public DocxPainter addToc(P title) {
        try {
            this.tocGenerator = new TocGenerator(wpk);
            SdtBlock sdtBlock = tocGenerator.generateToc(this.wpk.getMainDocumentPart().getContent().size(),
                    " TOC \\o \"1-3\" \\h \\u ", true);
            sdtBlock.getSdtContent().getContent().set(0, title);
        } catch (TocException e) {
            this.exceptionCatch(e);
        }
        return this;
    }

    public DocxPainter exception(Consumer<Exception> exceptionCapture) {
        this.exceptionCapture = exceptionCapture;
        return this;
    }

    public DocxPainter save(File target) {
        try {
            File folder = target.getParentFile();
            if (folder != null && !folder.exists()) {
                folder.mkdirs();
            }
            return this.save(Files.newOutputStream(target.toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public DocxPainter save(OutputStream out) {
        try {
            if (this.tocGenerator != null) {
                this.tocGenerator.updateToc(false);
            }
            Docx4J.save(wpk, out);
        } catch (Docx4JException e) {
            this.exceptionCatch(e);
        }
        return this;
    }

    private void exceptionCatch(Exception e) {
        this.exceptionCapture.accept(e);
    }

}
