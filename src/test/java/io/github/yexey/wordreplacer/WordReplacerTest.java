package io.github.yexey.wordreplacer;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

@Slf4j
class WordReplacerTest {

    @TempDir
    Path tempDir;

    @Test
    public void testReplacements() throws Exception {
        // Testdatei-Pfade vorbereiten und Ressourcen laden
        Path templatePath = tempDir.resolve("template.docx");
        Path outputPath = tempDir.resolve("output.docx");

        // Template aus src/test/resources in temporäres Verzeichnis kopieren
        try (InputStream resourceStream = getClass().getClassLoader().getResourceAsStream("template.docx")) {
            assertNotNull(resourceStream, "template.docx konnte nicht in src/test/resources gefunden werden");
            Files.copy(resourceStream, templatePath, StandardCopyOption.REPLACE_EXISTING);
        }

        // 1. Führe die Ersetzungen durch
        performReplacements(templatePath.toString(), outputPath.toString());

        // 2. Öffne das erzeugte Dokument
        try (FileInputStream fis = new FileInputStream(outputPath.toFile())) {
            XWPFDocument outputDoc = new XWPFDocument(fis);

            // 3. Extrahiere den gesamten Text aus dem Output-Dokument
            String fullText = extractFullText(outputDoc);

            // 4. Prüfe die korrekten Ersetzungen
            assertTrue(fullText.contains("Hier ist mein Text John Doe"), "Name wurde nicht korrekt ersetzt");
            assertTrue(fullText.contains("EMAIL: john.doe@example.com"), "Email wurde nicht korrekt ersetzt");
            assertTrue(fullText.contains("(555) 123-4567"), "Telefonnummer wurde nicht korrekt ersetzt");
            assertTrue(fullText.contains("2023-03-03"), "Datum wurde nicht korrekt ersetzt");

            // 5. Prüfe, dass der NOTES-Paragraph entfernt wurde
            assertFalse(fullText.contains("{{NOTES}}"), "NOTES-Platzhalter wurde nicht entfernt");

            // 6. Prüfe, dass wichtige Texte erhalten blieben
            assertTrue(fullText.contains("This should not be removed1"), "Wichtiger Text 1 wurde fälschlicherweise entfernt");
            assertTrue(fullText.contains("This should not be removed 2"), "Wichtiger Text 2 wurde fälschlicherweise entfernt");

            // 7. Prüfe, dass keine anderen Platzhalter im Dokument verblieben sind
            assertFalse(fullText.contains("{{"), "Es sind noch Platzhalter im Dokument vorhanden");

            // 8. Zähle die Tabellenzellen mit "(555) 123-4567" (sollten 3 sein)
            int phoneCount = countPhoneNumberOccurrences(outputDoc);
            assertEquals(3, phoneCount, "Telefonnummer sollte in genau 3 Tabellenzellen vorkommen");

            log.info("Alle Validierungsprüfungen bestanden");
        }
    }

    /**
     * Führt die Ersetzungen im Word-Dokument durch
     */
    private void performReplacements(String templatePath, String outputPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(templatePath)) {
            XWPFDocument document = new XWPFDocument(fis);

            WordReplacer replacer = new WordReplacer(document);

            // Einzelne Ersetzung
            replacer.replace("{{NAME}}", "John Doe");

            // Batch-Ersetzungen
            Map<String, String> replacements = new HashMap<>();
            replacements.put("{{EMAIL}}", "john.doe@example.com");
            replacements.put("{{PHONE}}", "(555) 123-4567");
            replacements.put("{{DATE}}", "2023-03-03");
            replacer.replace(replacements);

            // Standardwert-Ersetzung
            replacer.replaceOrDefault("{{ADDRESS}}", null, "No address provided");

            // Entferne Paragraphen
            replacer.removeParagraph("{{NOTES}}");

            // Protokolliere Statistiken
            log.info("Erfolgreiche Ersetzungen: " + replacer.getTracker().getReplacementCounts());
            log.info("Fehlgeschlagene Ersetzungen: " + replacer.getTracker().getFailedReplacements());

            // Speichere das Dokument
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }

    /**
     * Extrahiert den gesamten Text aus dem Word-Dokument
     */
    private String extractFullText(XWPFDocument document) {
        StringBuilder text = new StringBuilder();

        // Text aus Paragraphen extrahieren
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            text.append(paragraph.getText()).append("\n");
        }

        // Text aus Tabellen extrahieren
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        text.append(paragraph.getText()).append("\n");
                    }
                }
            }
        }

        return text.toString();
    }

    /**
     * Zählt die Vorkommen der Telefonnummer in Tabellenzellen
     */
    private int countPhoneNumberOccurrences(XWPFDocument document) {
        int count = 0;

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    String cellText = cell.getText();
                    if (cellText.contains("(555) 123-4567")) {
                        count++;
                    }
                }
            }
        }

        return count;
    }

    /**
     * Dieser Test vergleicht direkt die template.docx und output.docx Dateien
     */
    @Test
    public void testCompareTemplateAndOutputDocuments() throws Exception {
        // Testdatei-Pfade vorbereiten und Ressourcen laden
        Path templatePath = tempDir.resolve("template.docx");
        Path outputPath = tempDir.resolve("output.docx");

        // Template aus src/test/resources in temporäres Verzeichnis kopieren
        try (InputStream resourceStream = getClass().getClassLoader().getResourceAsStream("template.docx")) {
            assertNotNull(resourceStream, "template.docx konnte nicht in src/test/resources gefunden werden");
            Files.copy(resourceStream, templatePath, StandardCopyOption.REPLACE_EXISTING);
        }

        // Erzeuge die Output-Datei mit den Ersetzungen
        performReplacements(templatePath.toString(), outputPath.toString());

        // Lade Template und Output-Dokument
        try (FileInputStream templateFis = new FileInputStream(templatePath.toFile());
             FileInputStream outputFis = new FileInputStream(outputPath.toFile())) {

            XWPFDocument templateDoc = new XWPFDocument(templateFis);
            XWPFDocument outputDoc = new XWPFDocument(outputFis);

            // Sammle alle Platzhalter aus Template
            List<String> placeholders = findAllPlaceholders(templateDoc);
            log.info("Gefundene Platzhalter im Template: " + placeholders);

            // Prüfe, dass kein Platzhalter mehr im Output existiert
            for (String placeholder : placeholders) {
                String outputText = extractFullText(outputDoc);
                assertFalse(outputText.contains(placeholder),
                        "Platzhalter '" + placeholder + "' sollte im Output-Dokument ersetzt sein");
            }

            // Prüfe die Anzahl der Paragraphen (sollte um 1 reduziert sein wegen {{NOTES}})
            assertEquals(templateDoc.getParagraphs().size() - 1, outputDoc.getParagraphs().size(),
                    "Nach Entfernung des NOTES-Paragraphen sollte die Anzahl der Paragraphen um 1 reduziert sein");
        }
    }

    /**
     * Findet alle Platzhalter im Format {{XXX}} im Dokument
     */
    private List<String> findAllPlaceholders(XWPFDocument document) {
        List<String> placeholders = new ArrayList<>();
        String fullText = extractFullText(document);

        int startIndex = 0;
        while (true) {
            int openBraceIndex = fullText.indexOf("{{", startIndex);
            if (openBraceIndex == -1) break;

            int closeBraceIndex = fullText.indexOf("}}", openBraceIndex);
            if (closeBraceIndex == -1) break;

            String placeholder = fullText.substring(openBraceIndex, closeBraceIndex + 2);
            placeholders.add(placeholder);
            startIndex = closeBraceIndex + 2;
        }

        return placeholders;
    }
}