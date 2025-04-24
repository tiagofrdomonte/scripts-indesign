//#target indesign

// Função para criar a janela de entrada
function criarJanela() {
    var janela = new Window('dialog', 'Produção de anúncios - OPEC');

    var scriptFolder = Folder(app.activeScript).parent; // Pasta do script
    var logoPath = File(scriptFolder + "/logo.png")

    if (logoPath.exists) {
        var logoImage = janela.add('image', undefined, logoPath);
        logoImage.size = [150, 59]; // Ajuste o tamanho conforme necessário
        logoImage.alignment = 'center'; // Centralizar o logotipo
    }
  
    // Primeira opção: Escolher entre noticiário ou classificados
    var grupoTipo = janela.add('group');
    grupoTipo.add('statictext', undefined, 'Tipo de colunagem:');
    var tipoDropdown = grupoTipo.add('dropdownlist', undefined, ['Classificados', 'Noticiário O Dia', 'Noticiário MH']);
    tipoDropdown.selection = 0; // Padrão: Noticiário

    // Grupo de entrada para número de colunas
    var grupoColunas = janela.add('group');
    grupoColunas.add('statictext', undefined, 'Colunas:');
    var colunasDropdown = grupoColunas.add('dropdownlist', undefined, ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']); // Opções de colunas
    colunasDropdown.selection = 1; // Padrão: 1 coluna

    // Grupo de entrada para altura
    var grupoAltura = janela.add('group');
    grupoAltura.add('statictext', undefined, 'Altura (cm):');
    var alturaInput = grupoAltura.add('dropdownlist'); // Opções de altura

    // Preencher o dropdown com alturas de 1 a 52 cm
    for (var i = 2; i <= 52; i++) {
        alturaInput.add('item', i.toString());
    }
    alturaInput.selection = 0; // Padrão: 50 cm

    // Grupo para o título
    var grupoTitulo = janela.add('group');
    grupoTitulo.add('statictext', undefined, 'Título:');
    var tituloInput = grupoTitulo.add('edittext', undefined, '', {multiline: true, scrolling: true}); // Título padrão
    tituloInput.characters = 40;
    tituloInput.preferredSize = [400, 100];

    // Grupo para o corpo do texto
    var grupoTexto = janela.add('group');
    grupoTexto.add('statictext', undefined, 'Texto:');
    var textoInput = grupoTexto.add('edittext', undefined, '', {multiline: true, scrolling: true}); // Texto padrão
    textoInput.characters = 40;
    textoInput.preferredSize = [400, 150]; // Aumentar o campo de texto para 150px de altura e 300px de largura

    // Botões OK e Cancelar
    var botoes = janela.add('group');
    botoes.alignment = 'center';    function ajustarParagrafos(textFrame) {
        if (!textFrame || !(textFrame instanceof TextFrame)) {
            throw new Error("Um textFrame válido deve ser fornecido.");
        }
    
        var story = textFrame.parentStory;
    
        // Configurações iniciais
        var minPointSize = 6; // Tamanho mínimo da fonte
        var minTracking = -40; // Tracking mínimo
        var minHorizontalScale = 80; // Escala horizontal mínima
        var insetSpacingLimit = 0.5; // Limite de inset spacing
    
        // Loop para ajustar gradativamente
        for (var i = 0; i < 100; i++) { // Limite de 100 iterações para evitar loops infinitos
            var adjusted = false;
    
            // Ajustar o tamanho da fonte e a entrelinha
            for (var p = 0; p < story.paragraphs.length; p++) {
                var paragraph = story.paragraphs[p];
                if (paragraph.pointSize > minPointSize) {
                    paragraph.pointSize -= 0.5; // Reduzir 0.5pt
                    paragraph.leading = paragraph.pointSize; // Ajustar a entrelinha para o mesmo valor do corpo de letra
                    adjusted = true;
                }
            }
    
            // Ajustar o tracking
            if (story.texts[0].tracking > minTracking) {
                story.texts[0].tracking -= 10; // Reduzir tracking em -10
                adjusted = true;
            }
    
            // Ajustar a escala horizontal
            if (story.texts[0].horizontalScale > minHorizontalScale) {
                story.texts[0].horizontalScale -= 5; // Reduzir escala horizontal em 5
                adjusted = true;
            }
    
            // Verificar se o excesso de texto foi resolvido
            if (!textFrame.overflows) {
                return; // Excesso de texto resolvido, sair da função
            }
    
            // Se nenhuma alteração foi feita e o texto ainda está em excesso, ajustar o inset spacing
            if (!adjusted) {
                var insets = textFrame.textFramePreferences.insetSpacing;
                if (insets[0] > insetSpacingLimit || insets[1] > insetSpacingLimit || insets[2] > insetSpacingLimit || insets[3] > insetSpacingLimit) {
                    textFrame.textFramePreferences.insetSpacing = [
                        Math.max(insets[0] - 0.5, insetSpacingLimit),
                        Math.max(insets[1] - 0.5, insetSpacingLimit),
                        Math.max(insets[2] - 0.5, insetSpacingLimit),
                        Math.max(insets[3] - 0.5, insetSpacingLimit)
                    ];
                    adjusted = true;
                }
            }
    
            // Verificar novamente se o excesso de texto foi resolvido após ajustar o inset spacing
            if (!textFrame.overflows) {
                return; // Excesso de texto resolvido, sair da função
            }
    
            // Se nenhuma alteração foi feita, sair do loop
            if (!adjusted) {
                break;
            }
        }
    
        // Caso o excesso de texto não seja resolvido após todas as tentativas
        alert("Não foi possível ajustar o excesso de texto.");
    }
    var okButton = botoes.add('button', undefined, 'OK', {name: 'ok'});
    var cancelButton = botoes.add('button', undefined, 'Cancelar', {name: 'cancel'});

    // Exibir a janela e obter os dados do usuário
    if (janela.show() == 1) { // Se o usuário clicar em "OK"
        var largura = calcularLargura(tipoDropdown.selection.text, colunasDropdown.selection.text);
        var altura = parseInt(alturaInput.selection.text) * 10; // Converter cm para mm
        
        return {
            tipo: tipoDropdown.selection.text,
            largura: largura,
            altura: altura,
            titulo: tituloInput.text,
            texto: textoInput.text
        };
    } else {
        return null;
    }
}

// Função para substituir quebras de linha (\n) no corpo do texto, mas não no título
function substituirQuebraNoTexto(texto) {
    return texto.replace(/\n/g, ' ');
}

// Função para calcular a largura com base no tipo e número de colunas
function calcularLargura(tipo, colunas) {
    var larguraPorColuna;
    switch (tipo) {
        case 'Noticiário O Dia':
            larguraPorColuna = [46, 96, 146, 196, 246, 297, 362, 412, 462, 512]; // Larguras para Noticiário O Dia
            break;
        case 'Classificados':
            larguraPorColuna = [27, 57, 87, 117, 147, 177, 207, 237, 267, 297]; // Larguras para Classificados
            break;
        case 'Noticiário MH':
            larguraPorColuna = [47, 99, 151, 203, 255, 322, 374, 426, 478, 530]; // Larguras para Noticiário MH
            break;
        default:
            break;
    }
    var colunasIndex = parseInt(colunas) - 1; // O índice das colunas começa em 0
    return larguraPorColuna[colunasIndex]; // Retorna a largura correspondente
}

// Executar a função para criar a janela e obter os dados do usuário
var dados = criarJanela();
if (dados == null) {
    exit(); // Cancelar a execução do script se o usuário clicar em "Cancelar"
}

// Criar um novo documento com as dimensões inseridas pelo usuário
var doc = app.documents.add({
    documentPreferences: {
        pageWidth: dados.largura,
        pageHeight: dados.altura,
        facingPages: false,
        pagesPerDocument: 1,
        marginsPreferences: {
            top: 0,
            bottom: 0,
            left: 0,
            right: 0
        },
        columnCount: 1,
        columnGutter: 12,
        intent: DocumentIntentOptions.PRINT_INTENT
    }
});

// Acessar a primeira página do documento
var page = doc.pages.item(0);

// Zerar as margens da página
with (page.marginPreferences) {
    top = 0;
    bottom = 0;
    left = 0;
    right = 0;
}

// Criar um quadro de texto que ocupa todo o tamanho do documento
var textFrame = page.textFrames.add();
textFrame.geometricBounds = [0, 0, dados.altura, dados.largura];
textFrame.strokeWeight = 1;
textFrame.strokeAlignment = StrokeAlignment.INSIDE_ALIGNMENT;
textFrame.textFramePreferences.verticalJustification = VerticalJustification.JUSTIFY_ALIGN;
textFrame.textFramePreferences.insetSpacing = [1, 1, 1, 1];

// Código para adicionar o título e o corpo do texto
if (dados.titulo === '') {
    // Substituir quebras de linha no corpo do texto
    var textoModificado = substituirQuebraNoTexto(dados.texto);
    textFrame.contents = textoModificado;

    // Obter os parágrafos do texto
    var paragraphs = textFrame.paragraphs;

    // Aplicar formatação ao corpo do texto
    for (var i = 0; i < paragraphs.length; i++) {
        paragraphs[i].appliedFont = "Arial";
        paragraphs[i].pointSize = 7;
        paragraphs[i].justification = Justification.LEFT_JUSTIFIED;
        paragraphs[i].fontStyle = "Regular";        
    }

} else {
    // Substituir quebras de linha no corpo do texto, mantendo o título intacto
    var textoModificado = substituirQuebraNoTexto(dados.texto);
    textFrame.contents = dados.titulo + "\r" + textoModificado;

    var paragraphs = textFrame.paragraphs;

    if (paragraphs.length > 0) {
        // Formatar o título
        paragraphs[0].appliedFont = "Arial";
        paragraphs[0].pointSize = 8;
        paragraphs[0].justification = Justification.CENTER_ALIGN;
        paragraphs[0].fontStyle = "Bold";
        paragraphs[0].capitalization = Capitalization.ALL_CAPS; // Define o título como maiúsculo
    }

    // Formatar o corpo do texto
    for (var i = 1; i < paragraphs.length; i++) {
        paragraphs[i].appliedFont = "Arial";
        paragraphs[i].pointSize = 7;
        paragraphs[i].justification = Justification.LEFT_JUSTIFIED;
        paragraphs[i].fontStyle = "Regular";
    }
}

// Função de ajuste de excesso de texto
function ajustarParagrafos(textFrame) {
    if (!textFrame || !(textFrame instanceof TextFrame)) {
        throw new Error("Um textFrame válido deve ser fornecido.");
    }

    if (!textFrame.overflows) {
        return; // Não há excesso de texto, sair da função
    }

    var story = textFrame.parentStory;

    // Configurações iniciais
    var minPointSize = 6; // Tamanho mínimo da fonte
    var minTracking = -40; // Tracking mínimo
    var minHorizontalScale = 80; // Escala horizontal mínima
    var insetSpacingLimit = 0.5; // Limite de inset spacing

    // Loop para ajustar gradativamente
    for (var i = 0; i < 100; i++) { // Limite de 100 iterações para evitar loops infinitos
        var adjusted = false;

        // Ajustar o tamanho da fonte e a entrelinha
        for (var p = 0; p < story.paragraphs.length; p++) {
            var paragraph = story.paragraphs[p];
            if (paragraph.pointSize > minPointSize) {
                paragraph.pointSize -= 0.5; // Reduzir 0.5pt
                paragraph.leading = paragraph.pointSize; // Ajustar a entrelinha para o mesmo valor do corpo de letra
                adjusted = true;
            }
        }

        // Ajustar o tracking
        if (story.texts[0].tracking > minTracking) {
            story.texts[0].tracking -= 10; // Reduzir tracking em -10
            adjusted = true;
        }

        // Ajustar a escala horizontal
        if (story.texts[0].horizontalScale > minHorizontalScale) {
            story.texts[0].horizontalScale -= 5; // Reduzir escala horizontal em 5
            adjusted = true;
        }

        // Verificar se o excesso de texto foi resolvido
        if (!textFrame.overflows) {
            return; // Excesso de texto resolvido, sair da função
        }

        // Se nenhuma alteração foi feita e o texto ainda está em excesso, ajustar o inset spacing
        if (!adjusted) {
            var insets = textFrame.textFramePreferences.insetSpacing;
            if (insets[0] > insetSpacingLimit || insets[1] > insetSpacingLimit || insets[2] > insetSpacingLimit || insets[3] > insetSpacingLimit) {
                textFrame.textFramePreferences.insetSpacing = [
                    Math.max(insets[0] - 0.5, insetSpacingLimit),
                    Math.max(insets[1] - 0.5, insetSpacingLimit),
                    Math.max(insets[2] - 0.5, insetSpacingLimit),
                    Math.max(insets[3] - 0.5, insetSpacingLimit)
                ];
                adjusted = true;
            }
        }

        // Verificar novamente se o excesso de texto foi resolvido após ajustar o inset spacing
        if (!textFrame.overflows) {
            return; // Excesso de texto resolvido, sair da função
        }

        // Se nenhuma alteração foi feita, sair do loop
        if (!adjusted) {
            break;
        }
    }

    // Caso o excesso de texto não seja resolvido após todas as tentativas
    alert("Não foi possível ajustar o excesso de texto.");
}

ajustarParagrafos(textFrame);