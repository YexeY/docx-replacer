module org.yexey.wordreplacer {
    requires org.apache.poi.ooxml;
    requires org.slf4j;
    requires static lombok;  // 'static' weil lombok nur zur Kompilierzeit benötigt wird

    // Füge diese Zeile hinzu
    requires org.apache.commons.lang3;

    // Exportierte (öffentliche) Pakete
    exports org.yexey.wordreplacer;
}