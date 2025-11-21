//#target indesign

// ======================================================================
// JANELA PRINCIPAL – PRODUÇÃO DE ANÚNCIOS (ESTILO INDESIGN CLÁSSICO)
// ======================================================================

try {

    function decodeHtml(str) {
        if (!str) return "";

        var map = {
            "&amp;": "&",
            "&lt;": "<",
            "&gt;": ">",
            "&quot;": "\"",
            "&#39;": "'",
            "&apos;": "'",
            "&nbsp;": " "
        };

        // substitui entidades nomeadas
        for (var key in map) {
            str = str.split(key).join(map[key]);
        }

        // trata entidades numéricas (&#225; = á, etc.)
        str = str.replace(/&#(\d+);/g, function (match, dec) {
            return String.fromCharCode(dec);
        });

        // trata entidades hexadecimais (&#xE1; = á)
        str = str.replace(/&#x([0-9A-Fa-f]+);/g, function (match, hex) {
            return String.fromCharCode(parseInt(hex, 16));
        });

        return str;
    }
    function criarJanela() {

        var janela = new Window('dialog', 'Produção de Anúncios - OPEC');
        janela.orientation = 'column';
        janela.alignChildren = ['fill', 'top'];
        janela.spacing = 15;
        janela.margins = 20;

        // ------------------------------------------------------------
        // LOGO (opcional)
        // ------------------------------------------------------------
        try {
            var scriptFolder = Folder(app.activeScript).parent;
            var logoPath = File(scriptFolder + "/logo.png");
            if (logoPath.exists) {
                var logo = janela.add("image", undefined, logoPath);
                logo.alignment = "center";
            }
        } catch (e) { }

        // ------------------------------------------------------------
        // PAINEL DE CONFIGURAÇÕES
        // ------------------------------------------------------------
        var pnlConfig = janela.add('panel', undefined, 'Configurações');
        pnlConfig.orientation = 'column';
        pnlConfig.alignChildren = ['left', 'center'];
        pnlConfig.spacing = 12;
        pnlConfig.margins = 15;

        // Grupo Formato
        var grpFormato = pnlConfig.add('group');
        grpFormato.orientation = 'row';
        grpFormato.add('statictext', undefined, 'Formato:').preferredSize = [80, -1];

        var formatos = ['Classificados', 'Noticiário O Dia', 'Noticiário MH'];
        var tipoDropdown = grpFormato.add('dropdownlist', undefined, formatos);
        tipoDropdown.selection = 0;
        tipoDropdown.preferredSize = [200, 25];

        // Grupo Colunas
        var grpColunas = pnlConfig.add('group');
        grpColunas.orientation = 'row';
        grpColunas.add('statictext', undefined, 'Colunas:').preferredSize = [80, -1];

        var btnMenosCol = grpColunas.add('button', undefined, '–'); // en dash
        btnMenosCol.preferredSize = [30, 25];

        var colunasInput = grpColunas.add('edittext', undefined, '2');
        colunasInput.characters = 3;
        colunasInput.justify = 'center';

        var btnMaisCol = grpColunas.add('button', undefined, '+');
        btnMaisCol.preferredSize = [30, 25];

        // Grupo Altura
        var grpAltura = pnlConfig.add('group');
        grpAltura.orientation = 'row';
        grpAltura.add('statictext', undefined, 'Altura (cm):').preferredSize = [80, -1];

        var btnMenosAlt = grpAltura.add('button', undefined, '–');
        btnMenosAlt.preferredSize = [30, 25];

        var alturaInput = grpAltura.add('edittext', undefined, '2');
        alturaInput.characters = 3;
        alturaInput.justify = 'center';

        var btnMaisAlt = grpAltura.add('button', undefined, '+');
        btnMaisAlt.preferredSize = [30, 25];

        // Lógica dos botões
        var minCol = 1, maxCol = 10;
        var minAlt = 2, maxAlt = 52;

        function atualizarValor(input, valor, min, max) {
            var v = parseInt(valor, 10);
            if (isNaN(v)) v = min;
            if (v < min) v = min;
            if (v > max) v = max;
            input.text = v;
        }

        btnMenosCol.onClick = function () { atualizarValor(colunasInput, parseInt(colunasInput.text) - 1, minCol, maxCol); };
        btnMaisCol.onClick = function () { atualizarValor(colunasInput, parseInt(colunasInput.text) + 1, minCol, maxCol); };
        colunasInput.onChange = function () { atualizarValor(colunasInput, colunasInput.text, minCol, maxCol); };

        btnMenosAlt.onClick = function () { atualizarValor(alturaInput, parseInt(alturaInput.text) - 1, minAlt, maxAlt); };
        btnMaisAlt.onClick = function () { atualizarValor(alturaInput, parseInt(alturaInput.text) + 1, minAlt, maxAlt); };
        alturaInput.onChange = function () { atualizarValor(alturaInput, alturaInput.text, minAlt, maxAlt); };


        // ------------------------------------------------------------
        // PAINEL DE CONTEÚDO
        // ------------------------------------------------------------
        var pnlConteudo = janela.add('panel', undefined, 'Conteúdo');
        pnlConteudo.orientation = 'column';
        pnlConteudo.alignChildren = ['fill', 'top'];
        pnlConteudo.spacing = 8;
        pnlConteudo.margins = 15;

        pnlConteudo.add('statictext', undefined, 'Título:');
        var tituloInput = pnlConteudo.add('edittext', undefined, '', { multiline: true });
        tituloInput.preferredSize = [-1, 50];

        pnlConteudo.add('statictext', undefined, 'Texto:');
        var textoInput = pnlConteudo.add('edittext', undefined, '', { multiline: true, scrolling: true });
        textoInput.preferredSize = [-1, 120];

        // ------------------------------------------------------------
        // OPÇÕES E BOTÕES
        // ------------------------------------------------------------
        var quadroCheckbox = janela.add('checkbox', undefined, 'Quadro de texto principal (balanço + de 1 página)');
        quadroCheckbox.value = false;

        var grupoBotoes = janela.add('group');
        grupoBotoes.orientation = 'row';
        grupoBotoes.alignment = 'center';
        grupoBotoes.spacing = 10;

        var btnOK = grupoBotoes.add('button', undefined, 'OK', { name: 'ok' });
        var btnCancelar = grupoBotoes.add('button', undefined, 'Cancelar', { name: 'cancel' });

        // ------------------------------------------------------------
        // RETORNO
        // ------------------------------------------------------------
        if (janela.show() == 1) {
            var colunasVal = parseInt(colunasInput.text, 10);
            if (isNaN(colunasVal)) colunasVal = 1;
            var largura_mm = calcularLargura(tipoDropdown.selection.text, colunasVal);

            return {
                tipo: tipoDropdown.selection.text,
                colunas: colunasVal,
                largura_mm: largura_mm,
                altura_mm: parseInt(alturaInput.text, 10) * 10, // cm -> mm
                titulo: decodeHtml(tituloInput.text),
                texto: decodeHtml(textoInput.text),
                quadro: quadroCheckbox.value
            };
        } else {
            return null;
        }
    }

    // Função para substituir quebras de linha (\n) no corpo do texto, mas não no título
    function substituirQuebraNoTexto(texto) {
        if (texto == null) return '';
        return texto.replace(/\n/g, ' ');
    }

    // Função para calcular a largura com base no tipo e número de colunas
    function calcularLargura(tipo, colunas) {
        var larguraPorColuna;
        switch (tipo) {
            case 'Noticiário O Dia':
                larguraPorColuna = [46, 96, 146, 196, 246, 297, 362, 412, 462, 512];
                break;
            case 'Classificados':
                larguraPorColuna = [27, 57, 87, 117, 147, 177, 207, 237, 267, 297];
                break;
            case 'Noticiário MH':
                larguraPorColuna = [47, 99, 151, 203, 255, 322, 374, 426, 478, 530];
                break;
            default:
                // fallback para classificado
                larguraPorColuna = [27, 57, 87, 117, 147, 177, 207, 237, 267, 297];
                break;
        }
        var colIndex = parseInt(colunas, 10) - 1;
        if (isNaN(colIndex) || colIndex < 0) colIndex = 0;
        if (colIndex >= larguraPorColuna.length) colIndex = larguraPorColuna.length - 1;
        return larguraPorColuna[colIndex];
    }

    // Função de ajuste de excesso de texto (uma única implementação)
    function ajustarParagrafos(textFrame) {
        if (!textFrame || typeof textFrame.overflows === 'undefined') {
            throw new Error("Um textFrame válido deve ser fornecido.");
        }

        if (!textFrame.overflows) {
            return; // Não há excesso de texto
        }

        var story;
        try {
            story = textFrame.parentStory;
        } catch (e) {
            alert("Erro ao acessar parentStory: " + e.message);
            return;
        }

        // Configurações iniciais
        var minPointSize = 6;
        var minTracking = -40;
        var minHorizontalScale = 80;
        var insetSpacingLimit = 0.5;

        for (var i = 0; i < 100; i++) {
            var adjusted = false;

            // Reduzir ponto e leading nos parágrafos
            try {
                for (var p = 0; p < story.paragraphs.length; p++) {
                    var paragraph = story.paragraphs[p];
                    if (paragraph.pointSize > minPointSize) {
                        paragraph.pointSize = Math.max(paragraph.pointSize - 0.5, minPointSize);
                        paragraph.leading = paragraph.pointSize;
                        adjusted = true;
                    }
                }
            } catch (e) {
                // Continue mesmo se houver erro de estilo em algum parágrafo
            }

            // Tracking
            try {
                if (story.texts[0].tracking > minTracking) {
                    story.texts[0].tracking = Math.max(story.texts[0].tracking - 10, minTracking);
                    adjusted = true;
                }
            } catch (e) { }

            // Horizontal scale
            try {
                if (story.texts[0].horizontalScale > minHorizontalScale) {
                    story.texts[0].horizontalScale = Math.max(story.texts[0].horizontalScale - 5, minHorizontalScale);
                    adjusted = true;
                }
            } catch (e) { }

            // Se não houver overflow, finaliza
            if (!textFrame.overflows) return;

            // Ajustar inset spacing se nada mais fez diferença
            if (!adjusted) {
                try {
                    var insets = textFrame.textFramePreferences.insetSpacing;
                    var newInsets = [
                        Math.max(insets[0] - 0.5, insetSpacingLimit),
                        Math.max(insets[1] - 0.5, insetSpacingLimit),
                        Math.max(insets[2] - 0.5, insetSpacingLimit),
                        Math.max(insets[3] - 0.5, insetSpacingLimit)
                    ];
                    // Atualiza apenas se algum for maior que o limite
                    if (insets[0] > insetSpacingLimit || insets[1] > insetSpacingLimit || insets[2] > insetSpacingLimit || insets[3] > insetSpacingLimit) {
                        textFrame.textFramePreferences.insetSpacing = newInsets;
                        adjusted = true;
                    }
                } catch (e) { }
            }

            if (!textFrame.overflows) return;

            if (!adjusted) break; // sem alterações possíveis, sai
        }

        if (textFrame.overflows) {
            alert("Não foi possível ajustar completamente o excesso de texto.");
        }
    }

    // ============ Execução principal ============
    // Executar a função para criar a janela e obter os dados do usuário
    var dados = criarJanela();
    if (dados == null) {
        exit(); // Cancelar a execução do script se o usuário clicar em "Cancelar"
    }

    var doc = app.documents.add({
        documentPreferences: {
            // passar UnitValue em mm para garantir que InDesign interprete corretamente
            pageWidth: dados.largura_mm,
            pageHeight: dados.altura_mm,
            facingPages: false,
            pagesPerDocument: 1,
            columnCount: 1,
            columnGutter: 12,
            intent: DocumentIntentOptions.PRINT_INTENT,
            createPrimaryTextFrame: dados.quadro // Ativar o quadro de texto principal se a opção estiver marcada
        }
    });


    // Acessar a primeira página do documento
    var page = doc.pages.item(0);
    var masterPage = doc.masterSpreads[0].pages[0];

    // Zerar as margens da página-mestre
    try {
        masterPage.marginPreferences.top = 0;
        masterPage.marginPreferences.bottom = 0;
        masterPage.marginPreferences.left = 0;
        masterPage.marginPreferences.right = 0;
    } catch (e) { }

    if (dados.quadro == true) {
        // Garantir que exista um textFrame na masterPage antes de acessá-lo
        try {
            if (masterPage.textFrames.length == 0) {
                masterPage.textFrames.add();
            }
            // Redimensionar o quadro de texto principal da página-mestre (usar pontos)
            masterPage.textFrames[0].geometricBounds = [0, 0, , dados.largura_mm];
        } catch (e) {
            // não bloquear a execução só por isso
        }
    }

    // Zerar as margens da página
    // Zerar as margens da página
    try {
        page.marginPreferences.top = 0;
        page.marginPreferences.bottom = 0;
        page.marginPreferences.left = 0;
        page.marginPreferences.right = 0;
    } catch (e) { }

    if (dados.quadro == false) { var textFrame = page.textFrames.add(); } else { var textFrame = masterPage.textFrames[0]; }

    // Ajustar geometricBounds para utilizar pontos (top, left, bottom, right)
    textFrame.geometricBounds = [0, 0, dados.altura_mm, dados.largura_mm];
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

    }

    else {
        // Substituir quebras de linha no corpo do texto, mantendo o título intacto
        var tituloModificado = dados.titulo.replace(/^\s*/gm, '');
        var textoModificado = substituirQuebraNoTexto(dados.texto);
        textFrame.contents = tituloModificado + "\r" + textoModificado;

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

    ajustarParagrafos(textFrame);

} catch (e) {
    alert("Erro no script: " + e.message);
}
