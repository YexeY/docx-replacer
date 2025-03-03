package org.yexey.wordreplacer;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

@Slf4j
class WordReplacerTest {

    @Test
    public void test() throws Exception {
        // Load document
        try (FileInputStream fis = new FileInputStream("template.docx")) {
            XWPFDocument document = new XWPFDocument(fis);

            // Create replacer with builder pattern for configuration
            WordReplacer replacer = new WordReplacer(document);

            // Simple replacement
            replacer.replace("{{NAME}}", "John Doe");

            // Batch replacement
            Map<String, String> replacements = new HashMap<>();
            replacements.put("{{EMAIL}}", "john.doe@example.com");
            replacements.put("{{PHONE}}", "(555) 123-4567");
            replacements.put("{{DATE}}", "2023-03-03");
            replacer.replace(replacements);

            // Replace with default
            replacer.replaceOrDefault("{{ADDRESS}}", null, "No address provided");

            // Replace or remove paragraph
            replacer.removeParagraph("{{NOTES}}");

            // Get statistics
            log.info("Successful replacements: " +
                    replacer.getTracker().getReplacementCounts());
            log.info("Failed replacements: " +
                    replacer.getTracker().getFailedReplacements());

            // Save document
            try (FileOutputStream fos = new FileOutputStream("output.docx")) {
                document.write(fos);
            }
        }
    }
}