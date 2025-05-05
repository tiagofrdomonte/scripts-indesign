var doc = app.activeDocument;
var tables = doc.stories.everyItem().tables.everyItem().getElements();

doc.stories.everyItem().paragraphs.everyItem().appliedFont = app.fonts.item("Arial");
doc.stories.everyItem().paragraphs.everyItem().pointSize = 7;
doc.stories.everyItem().paragraphs.everyItem().leading = 7;
var myColor = doc.colors.itemByName("Black");
doc.stories.everyItem().paragraphs.everyItem().fillColor = myColor;

doc.stories.everyItem().paragraphs.everyItem().leftIndent = 0;
doc.stories.everyItem().paragraphs.everyItem().rightIndent = 0;
doc.stories.everyItem().paragraphs.everyItem().spaceBefore = 0;
doc.stories.everyItem().paragraphs.everyItem().spaceAfter = 0;

doc.stories.everyItem().paragraphs.everyItem().justification = Justification.LEFT_JUSTIFIED;

for (var i = 0; i < tables.length; i++) {
    var table = tables[i];

    table.cells.everyItem().topEdgeStrokeWeight = 0.5;
    table.cells.everyItem().bottomEdgeStrokeWeight = 0.5;
    table.cells.everyItem().leftEdgeStrokeWeight = 0.5;
    table.cells.everyItem().rightEdgeStrokeWeight = 0.5;

    table.cells.everyItem().topEdgeStrokeColor = "Black";
    table.cells.everyItem().bottomEdgeStrokeColor = "Black";
    table.cells.everyItem().leftEdgeStrokeColor = "Black";
    table.cells.everyItem().rightEdgeStrokeColor = "Black";

    table.cells.everyItem().paragraphs.everyItem().firstLineIndent = 0;
    table.cells.everyItem().paragraphs.everyItem().leftIndent = 0;
    table.cells.everyItem().paragraphs.everyItem().rightIndent = 0;
    table.cells.everyItem().paragraphs.everyItem().spaceBefore = 0;
    table.cells.everyItem().paragraphs.everyItem().spaceAfter = 0;

    for (var j = 0; j < table.rows.length; j++) {
        table.rows[j].minimumHeight = 1.058;
    }

    var textFrame = null;
    var parentObject = table.parent;

    while (parentObject && !(parentObject instanceof TextFrame)) {
        parentObject = parentObject.parent;
    }

    if (parentObject instanceof TextFrame) {
        textFrame = parentObject;
    }

    if (textFrame) {
        var columns = textFrame.textFramePreferences.textColumnCount;
        var gutter = textFrame.textFramePreferences.textColumnGutter;
        var totalWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
        var insets = textFrame.textFramePreferences.insetSpacing;
        var leftInset = insets[1];
        var rightInset = insets[3];
        var stroke = textFrame.strokeWeight || 0;
        var effectiveWidth = totalWidth - leftInset - rightInset - (stroke * 2);
        var columnWidth = (effectiveWidth - (gutter * (columns - 1))) / columns;
        
        var originalTotalWidth = 0;
        var colWidths = [];
        
        for (var k = 0; k < table.columnCount; k++) {
            colWidths.push(table.columns[k].width);
            originalTotalWidth += table.columns[k].width;
        }
        
        var scaleFactor = columnWidth / originalTotalWidth;
        
        for (var k = 0; k < table.columnCount; k++) {
            try {
                table.columns[k].width = colWidths[k] * scaleFactor + "mm"; // Adiciona unidade de medida
            } catch (e) {
                $.writeln("Erro ao ajustar a coluna " + k + ": " + e.message);
            }
        }
    } else {
        alert("A tabela " + (i + 1) + " não está dentro de um quadro de texto.");
    }

    table.cells.everyItem().topInset = 0.2;
    table.cells.everyItem().bottomInset = 0.2;
    table.cells.everyItem().leftInset = 0.2;
    table.cells.everyItem().rightInset = 0.2;

    table.cells.everyItem().texts[0].appliedFont = app.fonts.item("Arial");
    table.cells.everyItem().texts[0].pointSize = 7;
    table.cells.everyItem().texts[0].leading = 7;
    var myColor = doc.colors.itemByName("Black");
    table.cells.everyItem().texts.everyItem().fillColor = myColor;
}

alert("Todas as tabelas foram ajustadas proporcionalmente! ✅");
