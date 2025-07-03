function main() {



// Janela para configuração do usuário (layout melhorado)
var dlg = new Window('dialog', 'Formatar Tabelas');
dlg.orientation = 'column';
dlg.alignChildren = 'fill';

// Painel principal
var panel = dlg.add('panel', undefined, 'Parâmetros');
panel.orientation = 'column';
panel.alignChildren = 'left';
panel.margins = 15;

    var logoGroup = panel.add('group');
    logoGroup.orientation = 'row';
    logoGroup.alignment = 'center';

    try {
        var scriptFile = new File($.fileName);
        var scriptFolder = scriptFile.parent;
        var logoFile = File(scriptFolder + "/logo.png");
        if (logoFile.exists) {
            var logoImg = logoGroup.add('image', undefined, logoFile);
            logoImg.size = [150, 59]; // Ajuste o tamanho conforme necessário
        }
    } catch (e) {
        // Se não encontrar o logo, não faz nada
    }

    // Linha 1: Fonte e Entrelinha
    var row1 = panel.add('group');
    row1.orientation = 'row';
    row1.add('statictext', undefined, 'Tamanho da fonte:');
    var fontSizeInput = row1.add('edittext', undefined, '7');
    fontSizeInput.characters = 4;
    
    row1.add('statictext', undefined, 'Entrelinha:');
    var leadingInput = row1.add('edittext', undefined, '');
    leadingInput.characters = 4;

    // Linha 2: Colunas e Espaçamento interno
    var row2 = panel.add('group');
    row2.orientation = 'row';
    row2.add('statictext', undefined, 'Nº de colunas:');
    var columnsInput = row2.add('edittext', undefined, '1');
    columnsInput.characters = 3;
    
    row2.add('statictext', undefined, 'Espaço interno (mm):');
    var insetInput = row2.add('edittext', undefined, '1');
    insetInput.characters = 4;

    // Linha 3: Altura da linha
    var row3 = panel.add('group');
    row3.orientation = 'row';
    row3.add('statictext', undefined, 'Altura da linha (mm):');
    var rowHeightInput = row3.add('edittext', undefined, '1.058');
    rowHeightInput.characters = 5;

// Linha 4: Margens das células
var marginPanel = panel.add('panel', undefined, 'Margens da célula (mm)');
marginPanel.orientation = 'row';
marginPanel.alignChildren = 'top';

    var marginCol1 = marginPanel.add('group');
    marginCol1.orientation = 'column';
    marginCol1.add('statictext', undefined, 'Superior:');
    marginCol1.add('statictext', undefined, 'Inferior:');

    var marginCol2 = marginPanel.add('group');
    marginCol2.orientation = 'column';
    var topInput = marginCol2.add('edittext', undefined, '1');
    topInput.characters = 4;
    var bottomInput = marginCol2.add('edittext', undefined, '1');
    bottomInput.characters = 4;

    var marginCol3 = marginPanel.add('group');
    marginCol3.orientation = 'column';
    marginCol3.add('statictext', undefined, 'Esquerda:');
    marginCol3.add('statictext', undefined, 'Direita:');

    var marginCol4 = marginPanel.add('group');
    marginCol4.orientation = 'column';
    var leftInput = marginCol4.add('edittext', undefined, '1');
    leftInput.characters = 4;
    var rightInput = marginCol4.add('edittext', undefined, '1');
    rightInput.characters = 4;

// Botões
var btns = dlg.add('group');
btns.orientation = 'row';
btns.alignment = 'center';
var okBtn = btns.add('button', undefined, 'OK');
var cancelBtn = btns.add('button', undefined, 'Cancelar');

cancelBtn.onClick = function() {
    dlg.close(0);
};

if (dlg.show() != 1) exit();

// Pega valores do usuário
fontSizeInput.text = fontSizeInput.text.replace(",", ".");
var fontSize = parseFloat(fontSizeInput.text);
if (isNaN(fontSize)) {
    alert("Por favor, insira um tamanho de fonte válido.");
    main();
}

leadingInput.text = leadingInput.text.replace(",", ".");
var leading = leadingInput.text === "" ? undefined : parseFloat(leadingInput.text);

var numColumns = parseInt(columnsInput.text, 10) || 1;
var inset = parseFloat(insetInput.text) || 1;
var topInset = parseFloat(topInput.text) || 0;
var bottomInset = parseFloat(bottomInput.text) || 0;
var leftInset = parseFloat(leftInput.text) || 0;
var rightInset = parseFloat(rightInput.text) || 0;
var rowHeight = parseFloat(rowHeightInput.text) || 1.058;

var doc = app.activeDocument;
var tables = doc.stories.everyItem().tables.everyItem().getElements();

doc.stories.everyItem().paragraphs.everyItem().justification = Justification.LEFT_JUSTIFIED;
doc.stories.everyItem().paragraphs.everyItem().appliedFont = app.fonts.item("Arial");
doc.stories.everyItem().paragraphs.everyItem().pointSize = fontSize;
if (leading !== undefined) doc.stories.everyItem().paragraphs.everyItem().leading = leading;
var myColor = doc.colors.itemByName("Black");
doc.stories.everyItem().paragraphs.everyItem().fillColor = myColor;

doc.stories.everyItem().paragraphs.everyItem().leftIndent = 0;
doc.stories.everyItem().paragraphs.everyItem().rightIndent = 0;
doc.stories.everyItem().paragraphs.everyItem().spaceBefore = 0;
doc.stories.everyItem().paragraphs.everyItem().spaceAfter = 0;



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
        table.rows[j].minimumHeight = rowHeight;
    }

    var textFrame = null;
    var parentObject = table.parent;

    while (parentObject && !(parentObject instanceof TextFrame)) {
        parentObject = parentObject.parent;
    }

    if (parentObject instanceof TextFrame) {
        textFrame = parentObject;
    }

    // AJUSTE: altera o número de colunas do frame conforme escolha do usuário
    if (textFrame) {
        textFrame.textFramePreferences.textColumnCount = numColumns;
        textFrame.textFramePreferences.insetSpacing = inset;
        var columns = numColumns;
        var gutter = textFrame.textFramePreferences.textColumnGutter;
        var totalWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
        var insetsFrame = textFrame.textFramePreferences.insetSpacing;
        var leftInsetFrame = insetsFrame[1];
        var rightInsetFrame = insetsFrame[3];
        var stroke = textFrame.strokeWeight || 0;
        var effectiveWidth = totalWidth - leftInsetFrame - rightInsetFrame - (stroke * 2);
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
                table.columns[k].width = colWidths[k] * scaleFactor + "mm";
            } catch (e) {
                $.writeln("Erro ao ajustar a coluna " + k + ": " + e.message);
            }
        }
    } else {
        alert("A tabela " + (i + 1) + " não está dentro de um quadro de texto.");
    }

    table.cells.everyItem().topInset = topInset;
    table.cells.everyItem().bottomInset = bottomInset;
    table.cells.everyItem().leftInset = leftInset;
    table.cells.everyItem().rightInset = rightInset;

    table.cells.everyItem().texts[0].appliedFont = app.fonts.item("Arial");
    table.cells.everyItem().texts[0].pointSize = fontSize;
    if (leading !== undefined) table.cells.everyItem().texts[0].leading = leading;
    table.cells.everyItem().texts.everyItem().fillColor = myColor;
}

alert("Todas as tabelas foram ajustadas! ✅");

}

main();