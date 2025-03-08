package org.yexey.wordreplacer.internal.strategy.visitor.impl;

import org.apache.poi.xwpf.usermodel.*;
import org.yexey.wordreplacer.internal.strategy.visitor.DocumentElementVisitor;

import java.util.ArrayList;
import java.util.List;

import static org.yexey.wordreplacer.internal.utils.BookmarkUtils.containsBookmark;

/**
 * Visitor for removing elements containing bookmarks
 */
public class RemovalVisitor implements DocumentElementVisitor {
    private final String bookmark;
    private final XWPFDocument document;

    public RemovalVisitor(XWPFDocument document, String bookmark) {
        this.bookmark = bookmark;
        this.document = document;
    }

    @Override
    public void visitParagraph(XWPFParagraph paragraph) {
        if (containsBookmark(paragraph, bookmark)) {
            int pos = document.getPosOfParagraph(paragraph);
            if (pos >= 0) {
                document.removeBodyElement(pos);
            }
        }
    }

    @Override
    public void visitTable(XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            visitTableRow(row);
        }
    }

    @Override
    public void visitTableCell(XWPFTableCell cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            if (containsBookmark(paragraphs.get(i), bookmark)) {
                cell.removeParagraph(i);
                // Adjust index
                i--;
            }
        }

        // Process nested tables
        for (XWPFTable nestedTable : cell.getTables()) {
            visitTable(nestedTable);
        }
    }

    @Override
    public void visitTableRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            visitTableCell(cell);
        }
    }

    @Override
    public void visitHeader(XWPFHeader header) {
        List<XWPFParagraph> paragraphs = header.getParagraphs();
        List<XWPFParagraph> paragraphsToRemove = new ArrayList<>();

        for(var paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmark)) {
                paragraphsToRemove.add(paragraph);
            }
        }

        for(var paragraph : paragraphsToRemove) {
            header.removeParagraph(paragraph);
        }
    }

    @Override
    public void visitFooter(XWPFFooter footer) {
        List<XWPFParagraph> paragraphs = footer.getParagraphs();
        List<XWPFParagraph> paragraphsToRemove = new ArrayList<>();

        for(var paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmark)) {
                paragraphsToRemove.add(paragraph);
            }
        }

        for(var paragraph : paragraphsToRemove) {
            footer.removeParagraph(paragraph);
        }
    }
}