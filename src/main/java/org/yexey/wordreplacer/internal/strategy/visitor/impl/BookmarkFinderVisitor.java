package org.yexey.wordreplacer.internal.strategy.visitor.impl;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.*;
import org.yexey.wordreplacer.internal.strategy.visitor.DocumentElementVisitor;

import static org.yexey.wordreplacer.internal.utils.BookmarkUtils.containsBookmark;


/**
 * Visitor for finding bookmarks
 */
public class BookmarkFinderVisitor implements DocumentElementVisitor {
    private final String bookmark;
    @Getter
    private boolean found = false;

    public BookmarkFinderVisitor(String bookmark) {
        this.bookmark = bookmark;
    }

    @Override
    public void visitParagraph(XWPFParagraph paragraph) {
        if (containsBookmark(paragraph, bookmark)) {
            found = true;
        }
    }

    @Override
    public void visitTable(XWPFTable table) {
        if (!found) {
            for (XWPFTableRow row : table.getRows()) {
                visitTableRow(row);
                if (found) break;
            }
        }
    }

    @Override
    public void visitTableCell(XWPFTableCell cell) {
        if (!found) {
            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                visitParagraph(paragraph);
                if (found) break;
            }

            if (!found) {
                for (XWPFTable nestedTable : cell.getTables()) {
                    visitTable(nestedTable);
                    if (found) break;
                }
            }
        }
    }

    @Override
    public void visitTableRow(XWPFTableRow row) {
        if (!found) {
            for (XWPFTableCell cell : row.getTableCells()) {
                visitTableCell(cell);
                if (found) break;
            }
        }
    }

    @Override
    public void visitHeader(XWPFHeader header) {
        if (!found) {
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                visitParagraph(paragraph);
                if (found) break;
            }

            if (!found) {
                for (XWPFTable table : header.getTables()) {
                    visitTable(table);
                    if (found) break;
                }
            }
        }
    }

    @Override
    public void visitFooter(XWPFFooter footer) {
        if (!found) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                visitParagraph(paragraph);
                if (found) break;
            }

            if (!found) {
                for (XWPFTable table : footer.getTables()) {
                    visitTable(table);
                    if (found) break;
                }
            }
        }
    }
}