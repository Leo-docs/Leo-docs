function aplicarFormatacao(elementoTexto, inicio, fim, estilo, valor) {
  try {
    const estilosDeFormatacao = {
      titulo: {
        bold: true,
        color: "#000000",
        fontSize: 14,
      },
      paragrafo: {
        color: "#000000",
        fontSize: 12,
      },
      comentario: {
        italic: true,
        color: "#000000",
        fontSize: 12,
        fontFamily: "Arial",
      },
      citacao: {
        bold: false,
        italic: true,
        underline: false,
        color: "#000000",
        backgroundColor: null,
        fontSize: 12,
        fontFamily: "Georgia",
      },
      referencia: {
        bold: false,
        italic: false,
        underline: false,
        color: "#000000",
        backgroundColor: null,
        fontSize: 10,
        fontFamily: "Times New Roman",
      },
      link: {
        link: true,
        underline: true,
      },
    };

    const opcoesDeFormatacao = estilosDeFormatacao[estilo.toLowerCase()] || estilosDeFormatacao.paragrafo;

    if (opcoesDeFormatacao.bold) elementoTexto.setBold(inicio, fim, true);
    if (opcoesDeFormatacao.italic) elementoTexto.setItalic(inicio, fim, true);
    if (opcoesDeFormatacao.underline) elementoTexto.setUnderline(inicio, fim, true);
    if (opcoesDeFormatacao.color) elementoTexto.setForegroundColor(inicio, fim, opcoesDeFormatacao.color);
    if (opcoesDeFormatacao.backgroundColor) elementoTexto.setBackgroundColor(inicio, fim, opcoesDeFormatacao.backgroundColor);
    if (opcoesDeFormatacao.fontSize) elementoTexto.setFontSize(inicio, fim, opcoesDeFormatacao.fontSize);
    if (opcoesDeFormatacao.fontFamily) elementoTexto.setFontFamily(inicio, fim, opcoesDeFormatacao.fontFamily);
    if (opcoesDeFormatacao.link) elementoTexto.setLinkUrl(inicio, fim, valor);
  } catch (erro) {
    erro.funcao = `aplicarFormatacao / ${erro.funcao || ""}`;
    erro.dadosAdicionais = {
      contexto: `Elemento ${elementoTexto || "-"}, estilo ${estilo || "-"}, valor ${valor || "-"}`,
    };
    throw erro;
  }
}
