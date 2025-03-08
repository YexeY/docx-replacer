package org.yexey.wordreplacer.internal.strategy.tracker.impl;

import org.yexey.wordreplacer.internal.strategy.tracker.ReplacementTracker;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Class for tracking replacement statistics
 */
public class SimpleReplacementTracker implements ReplacementTracker {
    private final Map<String, Integer> replacementCounts = new HashMap<>();

    @Override
    public void trackReplacement(String bookmark, String replacement, boolean success) {
        replacementCounts.put(bookmark, replacementCounts.getOrDefault(bookmark, 0) + (success ? 1 : 0));
    }

    @Override
    public Map<String, Integer> getReplacementCounts() {
        return new HashMap<>(replacementCounts);
    }

    @Override
    public List<String> getFailedReplacements() {
        return replacementCounts.entrySet().stream()
                .filter(elm -> elm.getValue() == 0)
                .map(Map.Entry::getKey)
                .toList();
    }
}