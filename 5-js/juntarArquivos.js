document.addEventListener('DOMContentLoaded', () => {
    const btnProcessar = document.getElementById('btnProcessar');
    const btnLimpar = document.getElementById('btnLimpar');
    const inputArquivos = document.getElementById('arquivos');
    const inputGrupo = document.getElementById('grupo');
    const inputNomeArquivo = document.getElementById('nomeArquivo');
    const inputNumeroInicial = document.getElementById('numeroInicial'); // NOVO
    const statusContainer = document.getElementById('statusContainer');
    const statusTexto = document.getElementById('statusTexto');
    const statusBadge = document.getElementById('statusBadge');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const resultadosContainer = document.getElementById('resultadosContainer');
    const resultadosGrid = document.getElementById('resultadosGrid');

    btnProcessar.addEventListener('click', processar);
    btnLimpar.addEventListener('click', limpar);

    async function processar() {
        const grupo = parseInt(inputGrupo.value);

        // Pega o valor como texto para preservar os zeros à esquerda (ex: "010")
        const numeroInicialStr = inputNumeroInicial.value.trim() || "1";
        const numeroInicial = parseInt(numeroInicialStr, 10);
        const digitos = numeroInicialStr.length; // Define quantos zeros usar baseado no que o usuário digitou

        if (!inputArquivos.files.length) {
            mostrarStatus('Selecione pelo menos um arquivo!', 'warn');
            return;
        }
        if (grupo < 1) {
            mostrarStatus('O grupo deve ser pelo menos 1!', 'warn');
            return;
        }
        if (isNaN(numeroInicial) || numeroInicial < 1) {
            mostrarStatus('O número inicial deve ser válido!', 'warn');
            return;
        }

        btnProcessar.disabled = true;
        statusContainer.classList.remove('hidden');
        progressContainer.style.display = 'block';
        resultadosContainer.classList.add('hidden');
        mostrarStatus('Processando arquivos...', 'warn');

        try {
            const htmls = [];
            const nomes = [];
            const totalArquivos = inputArquivos.files.length;

            for (let i = 0; i < totalArquivos; i++) {
                const file = inputArquivos.files[i];
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
                htmls.push(result.value);
                nomes.push(file.name);
                progressBar.style.width = ((i + 1) / totalArquivos * 50) + '%';
            }

            const grupos = [];
            for (let i = 0; i < htmls.length; i += grupo) {
                const grupoHtmls = htmls.slice(i, i + grupo);
                const grupoNomes = nomes.slice(i, i + grupo);

                let htmlCombinado = `
                    <html>
                    <head><meta charset="UTF-8"></head>
                    <body style="font-family: Arial, sans-serif;">
                `;
                grupoHtmls.forEach((html, idx) => {
                    htmlCombinado += `
                        <div style="margin-bottom: 30px;">
                            <p style="color: #666; font-size: 12px; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 15px;">
                                <em>Arquivo original: ${grupoNomes[idx]}</em>
                            </p>
                            ${html}
                        </div>
                    `;
                });
                htmlCombinado += `</body></html>`;
                grupos.push({ html: htmlCombinado, nomes: grupoNomes });
            }

            mostrarStatus('Criando arquivos finais...', 'warn');
            const zip = new JSZip();

            const nomeBase = (inputNomeArquivo.value.trim() || 'arquivo').replace(/[^a-zA-Z0-9_-]/g, '_');

            for (let i = 0; i < grupos.length; i++) {
                const docx = htmlDocx.asBlob(grupos[i].html);
                const numeroAtual = numeroInicial + i;

                // Formata o número com a quantidade de dígitos que o usuário definiu no início
                const numeroFormatado = String(numeroAtual).padStart(digitos, '0');

                zip.file(`${nomeBase}_${numeroFormatado}.docx`, docx);
                progressBar.style.width = (50 + ((i + 1) / grupos.length * 50)) + '%';
            }

            const blob = await zip.generateAsync({ type: 'blob' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${nomeBase}_juntos.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            mostrarStatus(`${grupos.length} arquivo(s) gerado(s) com sucesso!`, 'ok');
            mostrarResultados(grupos, nomeBase, numeroInicial, digitos);
        } catch (err) {
            console.error(err);
            mostrarStatus('Erro ao processar: ' + err.message, 'err');
        } finally {
            btnProcessar.disabled = false;
            setTimeout(() => { progressContainer.style.display = 'none'; }, 1000);
        }
    }

    function mostrarResultados(grupos, nomeBase, numeroInicial, digitos) {
        resultadosGrid.innerHTML = '';
        grupos.forEach((grupo, idx) => {
            const numeroAtual = numeroInicial + idx;
            const numeroFormatado = String(numeroAtual).padStart(digitos, '0');

            const card = document.createElement('div');
            card.className = 'cardResultado';
            card.innerHTML = `
                <div class="iconeCard"><i class="fas fa-file-word"></i></div>
                <h3>${nomeBase}_${numeroFormatado}.docx</h3>
                <p>${grupo.nomes.length} arquivo(s) combinado(s)</p>
            `;
            resultadosGrid.appendChild(card);
        });
        resultadosContainer.classList.remove('hidden');
    }

    function mostrarStatus(msg, tipo) {
        statusTexto.textContent = msg;
        statusBadge.className = 'badgeStatus';

        if (tipo === 'ok') {
            statusBadge.classList.add('badgeOk');
            statusBadge.textContent = 'Sucesso';
        } else if (tipo === 'err') {
            statusBadge.classList.add('badgeErr');
            statusBadge.textContent = 'Erro';
        } else {
            statusBadge.classList.add('badgeWarn');
            statusBadge.textContent = 'Processando';
        }
    }

    function limpar() {
        inputArquivos.value = '';
        inputGrupo.value = '5';
        inputNomeArquivo.value = 'arquivo';
        inputNumeroInicial.value = '1'; // Reseta o número inicial
        statusContainer.classList.add('hidden');
        resultadosContainer.classList.add('hidden');
        progressBar.style.width = '0%';
        mostrarStatus('Aguardando', 'warn');
    }
});