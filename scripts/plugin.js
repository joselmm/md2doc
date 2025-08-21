// scripts/plugin.js
// Plugin para ONLYOFFICE: MD -> HTML -> insert / download DOCX
// Requiere: markdown-it, html-docx-js, FileSaver

(function () {
  'use strict';
  var theFormulas = [];
  // markdown-it inicial + katex
  const mdParser = window.markdownit({ html: true, linkify: true })


  // elementos UI
  const ta = document.getElementById('md-input');
  const preview = document.getElementById('preview');
  const btnInsert = document.getElementById('btn-insert');
  const btnDownload = document.getElementById('btn-download');
  const btnPreviewToggle = document.getElementById('btn-preview-toggle');
  const inputFile = document.getElementById('file-md');

  // 👉 función para extraer fórmulas del markdown
  function extractFormulas(markdownText) {
    const formulas = [];
    // Busca $$...$$ o $...$ (inline)
    const regex = /\$\$[\s\S]*?\$\$|\$[^$]+\$/g;
    let match;
    while ((match = regex.exec(markdownText)) !== null) {
      formulas.push(match[0].trim()); // match[0] incluye los $ o $$
    }
    theFormulas = formulas;
    return formulas;
  }




  // renderiza preview desde el markdown actual
  function renderPreview() {
    const text = ta.value || '';
    if (!mdParser) {
      preview.innerHTML = '<em>No hay markdown-it cargado.</em>';
      return;
    }
    try {
      const html = mdParser.render(text);
      preview.innerHTML = html;


      // 👇 Renderiza las fórmulas con KaTeX (auto-render)
      if (window.renderMathInElement) {
        renderMathInElement(preview, {
          delimiters: [
            { left: "$$", right: "$$", display: true },
            { left: "$", right: "$", display: false }
          ],
          throwOnError: false
        });
      }
      // 👇 extraer fórmulas
      const formulas = extractFormulas(text);
      console.log("📐 Fórmulas encontradas:", formulas);

      return html;
    } catch (err) {
      preview.innerHTML = '<pre style="color:red;">Error al parsear Markdown: ' + err.message + '</pre>';
      return '';
    }
  }

  // Función para convertir estilos CSS a inline
  function aplicarSoloEstilosDefinidos(elementoPadre) {
    function obtenerEstilosDefinidos(elemento) {
      let estilosDefinidos = '';

      // Obtener estilos de hojas de estilo
      for (let hoja of document.styleSheets) {
        try {
          for (let regla of hoja.cssRules) {
            if (regla.type === CSSRule.STYLE_RULE) {
              if (elemento.matches(regla.selectorText)) {
                estilosDefinidos += regla.style.cssText;
              }
            }
          }
        } catch (e) { } // Ignorar hojas de otros dominios
      }

      return estilosDefinidos;
    }

    // Aplicar al padre
    elementoPadre.style.cssText = obtenerEstilosDefinidos(elementoPadre);

    // Aplicar a hijos
    const hijos = elementoPadre.querySelectorAll('*');
    hijos.forEach(hijo => {
      hijo.style.cssText = obtenerEstilosDefinidos(hijo);
    });
  }

  function limpiarFormulas() {
    theFormulas = theFormulas.map(f => {
      let text = String(f).trim();

      // saltos reales -> espacio
      text = text.replace(/\r/g, '').replace(/\s*\n\s*/g, ' ');

      // saltos LaTeX: \\  \\*  \\[...]
      text = text.replace(/\\\\(\[[^\]]*\])?\*?/g, ' ');

      // quitar delimitadores de math
      text = text.replace(/^\s*\${2}\s*|\s*\${2}\s*$/g, '')   // $$...$$
        .replace(/^\s*\$\s*|\s*\$\s*$/g, '')         // $...$
        .replace(/^\s*\\\(\s*|\s*\\\)\s*$/g, '')     // \(...\)
        .replace(/^\s*\\\[\s*|\s*\\\]\s*$/g, '');    // \[...\]

      // compactar
      text = text.replace(/\s{2,}/g, ' ').trim();
      return text;
    });
    debugger
  }


  // Inserta el HTML convertido en el documento (ONLYOFFICE)
  function insertIntoDocument() {
    var html = renderPreview();
    extractFormulas(html)
    theFormulas.forEach((formul, i) => {
      html = html.replace(formul, "{formula" + i + "}")
    })


    limpiarFormulas();
    if (!html) {
      alert('Preview vacío — nada que insertar.');
      return;
    }

    // html completo recomendado por html->docx libs: incluimos wrapper
    const wrapped = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + html + '</body></html>';

    if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod && window.Asc.plugin.callCommand) {
      // PasteHtml inserta HTML en el documento. (ONLYOFFICE Plugin API)
      window.Asc.plugin.executeMethod('PasteHtml', [wrapped], function () {
        // callback opcional
        console.info('HTML insertado en el documento.');
        // puedes cerrar la UI si el plugin es modal: window.Asc.plugin.executeCommand("close", "");


      });
      setTimeout(() => {
        Asc.scope.theFormulas = theFormulas;
        Asc.plugin.callCommand(() => {
          const doc = Api.GetDocument();
          const formulas = Asc.scope.theFormulas;

          formulas.forEach((formula, index) => {
            const placeholder = `{formula${index}}`;
            const matches = doc.Search(placeholder);

            if (matches.length > 0) {
              const rng = matches[0];
              rng.Select();         // Selecciona el placeholder
              rng.Delete();         // Lo elimina
              doc.AddMathEquation(formula, "latex"); // Inserta la ecuación
            }
          });
        }, false);
      }, 0);


    } else {
      // Entorno de prueba (fuera de ONLYOFFICE) — simplemente descarga el HTML
      const blob = new Blob([wrapped], { type: 'text/html;charset=utf-8' });
      saveAs(blob, 'preview.html');
    }
  }

  // Genera un .docx del HTML y lo descarga (cliente)
  async function downloadDocx() {
    renderPreview(); // o asegurarte de que KaTeX ya corrió
    aplicarSoloEstilosDefinidos(preview); // <-- aquí también (CORREGIDO)

    const html = preview.innerHTML;
    const fullHtml = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<style>body{font-family:Arial,Helvetica,sans-serif}</style>' +
      '</head><body>' + html + '</body></html>';

    try {
      // html-docx-js: asBlob
      // Nota: html-docx-js a veces añade altChunk — prueba el resultado en Word/LibreOffice.
      if (window.htmlDocx && typeof window.htmlDocx.asBlob === 'function') {
        const blob = window.htmlDocx.asBlob(fullHtml);
        saveAs(blob, 'document-from-md.docx');
      } else {
        // Fallback simple: guardar HTML si no está la librería
        const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
        saveAs(blob, 'document-from-md.html');
      }
    } catch (err) {
      console.error(err);
      alert('Error generando DOCX: ' + err.message);
    }
  }

  // arrastrar/soltar o abrir archivo .md
  inputFile.addEventListener('change', (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      ta.value = e.target.result || '';
      renderPreview();
    };
    reader.readAsText(f, 'utf-8');
  });

  // render en tiempo real (simple debounce)
  let tId = null;
  ta.addEventListener('input', () => {
    clearTimeout(tId);
    tId = setTimeout(renderPreview, 150);
  });

  // botones
  btnInsert.addEventListener('click', insertIntoDocument);
  btnDownload.addEventListener('click', downloadDocx);
  btnPreviewToggle.addEventListener('click', () => {
    preview.style.display = (preview.style.display === 'none') ? 'block' : 'none';
  });

  // soporte para arrastrar al textarea
  ta.addEventListener('dragover', (e) => { e.preventDefault(); });
  ta.addEventListener('drop', (e) => {
    e.preventDefault();
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!f.type || f.name.toLowerCase().endsWith('.md') || f.type.indexOf('text') !== -1) {
      const r = new FileReader();
      r.onload = (ev) => { ta.value = ev.target.result || ''; renderPreview(); };
      r.readAsText(f, 'utf-8');
    } else alert('Archivo no soportado (sube un .md o archivo de texto).');
  });

  // ONLYOFFICE plugin init hook (si estamos dentro del editor)
  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = function () {
      renderPreview();
      window.Asc.plugin.executeMethod("GetSelectedText", [], function (selectedText) {
        if (ta && typeof selectedText === "string") {
          ta.value = selectedText;
          renderPreview();
        }
      });
    };
  } else {
    // Si pruebas en navegador fuera de ONLYOFFICE
    window.addEventListener('load', renderPreview);
  }

})();