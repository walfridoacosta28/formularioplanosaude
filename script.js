document.addEventListener('DOMContentLoaded', () => {

  // checar se a biblioteca XLSX está disponível
  if (typeof XLSX === 'undefined') {
    console.error('Biblioteca XLSX não encontrada. Inclua: <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>');
    alert('Erro: biblioteca XLSX não carregada. Verifique inclusão no HTML.');
    return;
  }

  // referências dos botões/elementos (IDs do seu HTML)
  const btnAddDep = document.getElementById('btnAddDep');
  const btnGerar = document.getElementById('btnGerar');
  const btnLimpar = document.getElementById('btnLimpar');
  const tabelaDep = document.getElementById('tabelaDep');
  const inputNomeTitular = document.getElementById('nome');

  // função utilitária para escapar texto ao inserir em value
  function escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // adicionar dependente (mesma estrutura do seu HTML)
  function adicionarLinha() {
    const linha = tabelaDep.insertRow(-1);
    const idx = tabelaDep.rows.length - 1;
    const titularName = inputNomeTitular.value || '';
    linha.innerHTML = `
      <td>${String(idx).padStart(2,'0')}</td>
      <td><input type='text' value='${escapeHtml(titularName)}' readonly></td>
      <td><input required type='text' placeholder='Nome'></td>
      <td><input required type='text' placeholder='CPF'></td>
      <td><input type='text' placeholder='Identidade'></td>
      <td><input type='date' placeholder='Nasc.'></td>
      <td><input type='text' placeholder='Sexo'></td>
      <td><input type='text' placeholder='Valor'></td>
      <td><button type='button' class='btnDel'>Excluir</button></td>
    `;

    // anexar evento ao botão Excluir recém-criado
    const btnDel = linha.querySelector('.btnDel');
    if (btnDel) {
      btnDel.addEventListener('click', () => excluirDependente(btnDel));
    }
  }

  // excluir dependente (recebe o botão ou referência da linha)
  function excluirDependente(btnOrRow) {
    let row;
    if (btnOrRow instanceof HTMLButtonElement) {
      row = btnOrRow.closest('tr');
    } else if (btnOrRow instanceof HTMLTableRowElement) {
      row = btnOrRow;
    } else {
      return;
    }

    row.parentNode.removeChild(row);

    // reindexar coluna D
    const linhas = tabelaDep.querySelectorAll('tr');
    let contador = 1;
    for (let i = 1; i < linhas.length; i++) {
      const cel = linhas[i].cells[0];
      if (cel) cel.innerText = String(contador).padStart(2,'0');
      contador++;
    }
  }

  // atualizar automaticamente o nome do titular em todas as linhas de dependentes
  function atualizarTitularNosDependentes(novoNome) {
    const inputsTitular = tabelaDep.querySelectorAll('tr td:nth-child(2) input');
    inputsTitular.forEach(inp => inp.value = novoNome);
  }

  // gerar excel - agrupa titular e dependentes e baixa arquivo
  function gerarExcel() {
    // validação básica: se quiser, pode usar form.reportValidity() no seu form
    // coletar dados do titular (IDs baseados no seu HTML atual)
    const principal = {
      Nome: (document.getElementById('nome') && document.getElementById('nome').value) || '',
      Identidade: (document.getElementById('identidade') && document.getElementById('identidade').value) || '',
      CPF_CNPJ: (document.getElementById('cpf') && document.getElementById('cpf').value) || '',
      Endereco: (document.getElementById('endereco') && document.getElementById('endereco').value) || '',
      Bairro: (document.getElementById('bairro') && document.getElementById('bairro').value) || '',
      Cidade: (document.getElementById('cidade') && document.getElementById('cidade').value) || '',
      Estado: (document.getElementById('estado') && document.getElementById('estado').value) || '',
      CEP: (document.getElementById('cep') && document.getElementById('cep').value) || '',
      Telefone: (document.getElementById('telefone') && document.getElementById('telefone').value) || '',
      Nascimento: (document.getElementById('nascimento') && document.getElementById('nascimento').value) || '',
      Sexo: (document.getElementById('sexo') && document.getElementById('sexo').value) || '',
      Empresa: (document.getElementById('empresa') && document.getElementById('empresa').value) || '',
      Cargo: (document.getElementById('cargo') && document.getElementById('cargo').value) || '',
      CNPJ_Empresa: (document.getElementById('cnpj_empresa') && document.getElementById('cnpj_empresa').value) || '',
      Categoria_Plano: (document.getElementById('categoria_plano') && document.getElementById('categoria_plano').value) || '',
      Email: (document.getElementById('email') && document.getElementById('email').value) || '',
      Contrato: (document.getElementById('contrato') && document.getElementById('contrato').value) || '',
      Valor: (document.getElementById('valor') && document.getElementById('valor').value) || '',
      Adesao: (document.getElementById('adesao') && document.getElementById('adesao').value) || '',
      Vencimento: (document.getElementById('vencimento') && document.getElementById('vencimento').value) || ''
    };

    // coletar dependentes
    const dependentes = [];
    const linhas = tabelaDep.querySelectorAll('tr');
    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];
      const inputs = row.querySelectorAll('input');
      // mapeamento seguro: se inputs >=7, considera formato esperado
      if (inputs.length >= 7) {
        dependentes.push({
          D: row.cells[0] ? row.cells[0].innerText : String(i).padStart(2,'0'),
          Titular: inputs[0] ? inputs[0].value : '',
          Nome: inputs[1] ? inputs[1].value : '',
          CPF: inputs[2] ? inputs[2].value : '',
          Identidade: inputs[3] ? inputs[3].value : '',
          Nascimento: inputs[4] ? inputs[4].value : '',
          Sexo: inputs[5] ? inputs[5].value : '',
          Valor: inputs[6] ? inputs[6].value : ''
        });
      } else {
        // formato inesperado: tentar extrair por posição das células
        dependentes.push({
          D: row.cells[0] ? row.cells[0].innerText : String(i).padStart(2,'0'),
          Titular: row.cells[1] ? row.cells[1].innerText || '' : '',
          Nome: row.cells[2] ? (row.cells[2].querySelector('input') ? row.cells[2].querySelector('input').value : row.cells[2].innerText) : '',
          CPF: row.cells[3] ? (row.cells[3].querySelector('input') ? row.cells[3].querySelector('input').value : row.cells[3].innerText) : '',
          Identidade: row.cells[4] ? (row.cells[4].querySelector('input') ? row.cells[4].querySelector('input').value : row.cells[4].innerText) : '',
          Nascimento: row.cells[5] ? (row.cells[5].querySelector('input') ? row.cells[5].querySelector('input').value : row.cells[5].innerText) : '',
          Sexo: row.cells[6] ? (row.cells[6].querySelector('input') ? row.cells[6].querySelector('input').value : row.cells[6].innerText) : '',
          Valor: row.cells[7] ? (row.cells[7].querySelector('input') ? row.cells[7].querySelector('input').value : row.cells[7].innerText) : ''
        });
      }
    }

    // criar workbook e sheets via SheetJS
    const wb = XLSX.utils.book_new();
    const wsPrincipal = XLSX.utils.json_to_sheet([principal]);
    const wsDependentes = XLSX.utils.json_to_sheet(dependentes.length ? dependentes : []);
    XLSX.utils.book_append_sheet(wb, wsPrincipal, 'Principal');
    XLSX.utils.book_append_sheet(wb, wsDependentes, 'Dependentes');

    try {
      XLSX.writeFile(wb, 'cadastro_plano.xlsx');
      // se quiser exibir feedback:
      // alert('Arquivo gerado com sucesso: cadastro_plano.xlsx');
    } catch (err) {
      console.error('Erro ao gerar o Excel:', err);
      alert('Falha ao gerar o Excel. Veja console do navegador para detalhes.');
    }
  }

  // limpar formulário e dependentes (só executa ao clicar no botão Limpar)
  function limparFormulario() {
    const form = document.getElementById('formPrincipal');
    if (form) form.reset();
    // remover todas as linhas exceto cabeçalho
    while (tabelaDep.rows.length > 1) tabelaDep.deleteRow(1);
  }

  // atualiza titular automaticamente quando campo nome muda
  if (inputNomeTitular) {
    inputNomeTitular.addEventListener('input', (e) => {
      const novo = e.target.value || '';
      atualizarTitularNosDependentes(novo);
    });
  }

  // eventos dos botões
  if (btnAddDep) btnAddDep.addEventListener('click', adicionarLinha);
  if (btnGerar) btnGerar.addEventListener('click', gerarExcel);
  if (btnLimpar) btnLimpar.addEventListener('click', () => {
    if (confirm('Tem certeza que deseja limpar o formulário?')) limparFormulario();
  });

});
