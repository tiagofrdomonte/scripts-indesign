 function parseFloatBR(text) {
        return parseFloat(text.replace(",", "."));
    }

function main() {

   

// Janela para configuração do usuário (layout melhorado)
var dlg = new Window('dialog', 'Formatar Tabelas');
dlg.orientation = 'column';
dlg.alignChildren = 'fill';

var logoGroup = dlg.add('group');
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

    // Painel principal
var panel = dlg.add('panel', undefined, 'Parâmetros');
panel.orientation = 'column';
panel.alignChildren = 'left';
panel.margins = 20;

    // Linha 1: Fonte e Entrelinha
    var row1 = panel.add('group');
    row1.orientation = 'row';
    row1.add('statictext', undefined, 'Tamanho da fonte:');
    var fontSizeInput = row1.add('edittext', undefined, '7');
    fontSizeInput.characters = 3;

    row1.add('statictext', undefined, 'Entrelinha:');
    var leadingInput = row1.add('edittext', undefined, '');
    leadingInput.characters = 3;

    // Linha 2: Colunas e Espaçamento interno
    var row2 = panel.add('group');
    row2.orientation = 'row';
    row2.add('statictext', undefined, 'Nº de colunas:');
    var columnsInput = row2.add('edittext', undefined, '1');
    columnsInput.characters = 3;
    
    row2.add('statictext', undefined, 'Espaçamento interno:');
    var insetInput = row2.add('edittext', undefined, '1');
    insetInput.characters = 3;

    // Linha 3: Altura da linha
    var row3 = panel.add('group');
    row3.orientation = 'row';
    row3.add('statictext', undefined, 'Altura da linha (mm):');
    var rowHeightInput = row3.add('edittext', undefined, '1,058');
    rowHeightInput.characters = 4;

    // var marginCol5 = panel.add('group');
    // marginCol5.orientation = 'row';
    // var tracado = marginCol5.add('statictext', undefined, 'Espessura do Traçado (pt):');
    // var tracadoInput = marginCol5.add('edittext', undefined, '1');
    // tracadoInput.characters = 3;

// Linha 4: Margens das células
var marginPanel = panel.add('panel', undefined, 'Margens da célula (mm)');
marginPanel.alignment = 'fill';
marginPanel.orientation = 'column';
marginPanel.alignChildren = 'center';

    var marginLine1 = marginPanel.add('group');
    marginLine1.orientation = 'row';
    marginLine1.add('statictext', undefined, '↑:');
    var topInput = marginLine1.add('edittext', undefined, '1');
    topInput.characters = 4;

    marginLine1.add('statictext', undefined, '←:');
    var leftInput = marginLine1.add('edittext', undefined, '1');
    leftInput.characters = 4;
    
    var marginLine2 = marginPanel.add('group');
    marginLine2.orientation = 'row';
    marginLine2.add('statictext', undefined, '↓:');
    var bottomInput = marginLine2.add('edittext', undefined, '1');
    bottomInput.characters = 4;

    marginLine2.add('statictext', undefined, '→:');
    var rightInput = marginLine2.add('edittext', undefined, '1');
    rightInput.characters = 4;

// Botões
var btns = dlg.add('group');
btns.orientation = 'row';
btns.alignment = 'center';
var okBtn = btns.add('button', undefined, 'OK');
var cancelBtn = btns.add('button', undefined, 'Cancelar');

okBtn.onClick = function() {dlg.close(1);};

cancelBtn.onClick = function() {dlg.close(0);};

if (dlg.show() != 1) exit();


// Pega valores do usuário
var fontSize = parseFloatBR(fontSizeInput.text);
if (isNaN(fontSize)) {
    alert("Por favor, insira um tamanho de fonte válido.");
    main();
}

var leading = leadingInput.text === "" ? undefined : parseFloatBR(leadingInput.text);

var numColumns = parseInt(columnsInput.text, 10) || 1;
if (isNaN(numColumns) || numColumns < 1) {
    alert("Por favor, insira um número válido de colunas.");
    main();
}

var inset = parseFloatBR(insetInput.text) || 1;
var topInset = parseFloatBR(topInput.text) || 0;
var bottomInset = parseFloatBR(bottomInput.text) || 0;
var leftInset = parseFloatBR(leftInput.text) || 0;
var rightInset = parseFloatBR(rightInput.text) || 0;
var rowHeight = parseFloatBR(rowHeightInput.text) || 1.058;

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

if (app.documents.length === 0) {
    alert("Nenhum documento aberto.");
    exit();
}

var pages = doc.pages;
var removedCount = 0;

for (var i = pages.length - 1; i >= 0; i--) {
    var page = pages[i];
    var textFrames = page.textFrames;
    var allEmpty = true;

    if (textFrames.length === 0) {
        // Se não há quadros de texto, considere a página como vazia
        allEmpty = true;
    } else {
        for (var j = 0; j < textFrames.length; j++) {
            if (textFrames[j].contents !== "") {
                allEmpty = false;
                break;
            }
        }
    }

    if (allEmpty) {
        page.remove();
        removedCount++;
    }
}

alert("Todas as tabelas foram ajustadas! ✅\n" + "Páginas em branco removidas: " + removedCount + "\n\n");

}

main();