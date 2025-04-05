# WordReplacer

A Java library for replacing placeholders in Microsoft Word documents (.docx) using Apache POI.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

WordReplacer is a utility that helps you automate Microsoft Word document workflows by replacing placeholders with actual content. Whether you need to generate personalized documents, fill in templates, or create dynamic reports, WordReplacer provides a simple and flexible API to handle these tasks.

### Key Features

- Replace any placeholders with actual content in paragraphs, tables, headers, and footers
- Support for batch replacements to update multiple placeholders at once
- Conditional replacements with default values
- Remove entire paragraphs containing specified placeholders
- Track successful and failed replacements
- Works with all parts of Word documents including tables, headers, and footers
- Retain original foramtting
- It just works, where other libraries fail to robustly replace the placeholders

## Installation

### Maven

```xml
<!-- https://central.sonatype.com/artifact/io.github.yexey/word-replacer -->
<dependency>
    <groupId>io.github.yexey</groupId>
    <artifactId>word-replacer</artifactId>
    <version>1.0.0</version>
</dependency>
```

### Gradle

```groovy
implementation 'io.github.yexey:word-replacer:1.0.0'
```

## Usage

### Basic Example

```java
import io.github.yexey.wordreplacer.WordReplacer;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Example {
    public static void main(String[] args) throws Exception {
        // Open the Word document
        try (FileInputStream inputStream = new FileInputStream("template.docx")) {
            XWPFDocument document = new XWPFDocument(inputStream);
            
            // Create a WordReplacer instance
            WordReplacer replacer = new WordReplacer(document);
            
            // Replace a single placeholder
            replacer.replace("{{NAME}}", "John Doe");
            
            // Save the document
            try (FileOutputStream outputStream = new FileOutputStream("output.docx")) {
                document.write(outputStream);
            }
        }
    }
}
```

### Multiple Replacements

```java
import java.util.HashMap;
import java.util.Map;

// Create a map for batch replacements
Map<String, String> replacements = new HashMap<>();
replacements.put("{{NAME}}", "John Doe");
replacements.put("{{EMAIL}}", "john.doe@example.com");
replacements.put("{{PHONE}}", "(555) 123-4567");
replacements.put("{{DATE}}", "2023-03-03");

// Apply all replacements at once
replacer.replace(replacements);
```

### Conditional Replacements

```java
// Replace with default value if the replacement is null or empty
String address = null; // This might come from a database or user input
replacer.replaceOrDefault("{{ADDRESS}}", address, "No address provided");
```

### Removing Paragraphs

```java
// Remove entire paragraphs containing a specific placeholder
replacer.removeParagraph("{{NOTES}}");
```

### Checking for Placeholders

```java
// Check if a placeholder exists in the document
if (replacer.hasBookmark("{{OPTIONAL_SECTION}}")) {
    // Handle optional section...
}
```

### Tracking Replacement Statistics

```java
// Get statistics about successful replacements
System.out.println("Successful replacements: " + replacer.getTracker().getReplacementCounts());

// Get information about failed replacements
System.out.println("Failed replacements: " + replacer.getTracker().getFailedReplacements());
```

## Creating Templates

Templates should be regular Microsoft Word documents (.docx) with placeholders in the format `{{PLACEHOLDER}}`. For example:

- Name: {{NAME}}
- Email: {{EMAIL}}
- Phone: {{PHONE}}
- Date: {{DATE}}

Or even:

- Name: [NAME]
- Email: [EMAIL]
- Phone: <PHONE>
- Date: DATE

You can place these placeholders in:
- Regular paragraphs
- Table cells
- Headers and footers

## Advanced Usage

### Working with Optional Content

For sections that might be removed, use a consistent placeholder like `{{NOTES}}` for the entire paragraph and then use `removeParagraph()` to remove it when needed.

## Flexible Placeholder Format

There is no enforced format for placeholders. 
The examples in this documentation use {{PLACEHOLDER}}, but you can use any text pattern as a placeholder. 
You just need to ensure that you're looking for the same pattern in your code that you've used in your Word document. 
This gives you complete flexibility in how you design your templates.

## Requirements

- Java 8 or higher

## Building from Source

```bash
git clone https://github.com/yexey/word-replacer.git
cd word-replacer
mvn clean install
```

## Running Tests

```bash
mvn test
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request
