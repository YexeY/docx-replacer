package org.yexey.wordreplacer.strategy.tracker;

import java.util.List;
import java.util.Map;

/**
 * Interface for tracking replacement statistics
 */
public interface ReplacementTracker {
    void trackReplacement(String bookmark, String replacement, boolean success);
    Map<String, Integer> getReplacementCounts();
    List<String> getFailedReplacements();
}