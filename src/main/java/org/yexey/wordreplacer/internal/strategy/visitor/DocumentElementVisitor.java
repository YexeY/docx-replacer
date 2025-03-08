package org.yexey.wordreplacer.internal.strategy.visitor;

import org.apache.poi.xwpf.usermodel.*;

/**
 * Interface for visiting different document elements
 */
public interface DocumentElementVisitor {
    void visitParagraph(XWPFParagraph paragraph);
    void visitTable(XWPFTable table);
    void visitTableCell(XWPFTableCell cell);
    void visitTableRow(XWPFTableRow row);
    void visitHeader(XWPFHeader header);
    void visitFooter(XWPFFooter footer);
}