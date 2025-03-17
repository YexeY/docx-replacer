package io.github.yexey.wordreplacer;

import java.util.Map;
import java.util.Optional;

public interface WordReplacerIF {
    void replace(String bookmark, String replacement);

    void replace(Map<String, String> replacements);

    void replaceOrDefault(String bookmark, String replacement, String defaultText);

    void replace(String bookmark, Optional<String> replacement);

    void removeParagraph(String bookmark);

    boolean hasBookmark(String bookmark);
}
