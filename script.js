    function adicionarLinha() {
      const tabela = document.getElementById('tabelaDep');
      const linha = tabela.insertRow(-1);
      const idx = tabela.rows.length - 1;

      // Preencher titular com o nome atual do formulário (readonly)
      const titularName = document.getElementById('nome').value || '';

      linha.innerHTML = `
        <td>${String(idx).padStart(2,'0')}</td>
        <td><input type='text' value='${escapeHtml(titularName)}' readonly></td>
        <td><input required type='text'></td>
        <td><input required type='text'></td>
        <td><input type='text'></td>
        <td><input type='date'></td>
        <td><input type='text'></td>
        <td><input type='text'></td>
        <td><button type='button' class='btnDel' onclick='excluirDependente(this)'>Excluir</button></td>`;
    }

    // Função para excluir dependente
    function excluirDependente(btn) {
      const row = btn.parentNode.parentNode;
      row.parentNode.removeChild(row);

      // Reordenar numeração (coluna D)
      const linhas = document.querySelectorAll('#tabelaDep tr');
      let contador = 1;
      for (let i = 1; i < linhas.length; i++) {
        linhas[i].cells[0].innerText = String(contador).padStart(2,'0');
        contador++;
      }
    }

    // Função para gerar o Excel e depois limpar os campos
    function gerarExcel() {
      const form = document.getElementById('formPrincipal');
      if (!form.reportValidity()) return; // impede se houver campo inválido

      // Coletar dados do titular
      const principal = {
        Nome: document.getElementById('nome').value,
        Identidade: document.getElementById('identidade').value,
        CPF_CNPJ: document.getElementById('cpf').value,
        Endereco: document.getElementById('endereco').value,
        Bairro: document.getElementById('bairro').value,
        Cidade: document.getElementById('cidade').value,
        Estado: document.getElementById('estado').value,
        CEP: document.getElementById('cep').value,
        Telefone: document.getElementById('telefone').value,
        Nascimento: document.getElementById('nascimento').value,
        Sexo: document.getElementById('sexo').value,
        Empresa: document.getElementById('empresa').value,
        Cargo: document.getElementById('cargo').value,
        CNPJ_Empresa: document.getElementById('cnpj_empresa').value,
        Categoria_Plano: document.getElementById('categoria_plano').value,
        Email: document.getElementById('email').value,
        Contrato: document.getElementById('contrato').value,
        Valor: document.getElementById('valor').value,
        Adesao: document.getElementById('adesao').value,
        Vencimento: document.getElementById('vencimento').value
      };

      // Coletar dependentes
      const dependentes = [];
      const linhas = document.querySelectorAll('#tabelaDep tr');
      for (let i = 1; i < linhas.length; i++) {
        const inputs = linhas[i].querySelectorAll('input');
        dependentes.push({
          D: inputs[0] ? inputs[0].value || String(i).padStart(2,'0') : String(i).padStart(2,'0'),
          Titular: inputs[1] ? inputs[1].value : '',
          Nome: inputs[2] ? inputs[2].value : '',
          CPF: inputs[3] ? inputs[3].value : '',
          Identidade: inputs[4] ? inputs[4].value : '',
          Nascimento: inputs[5] ? inputs[5].value : '',
          Sexo: inputs[6] ? inputs[6].value : '',
          Valor: inputs[7] ? inputs[7].value : ''
        });
      }

      // Criar workbook e sheets
      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet([principal]);
      const ws2 = XLSX.utils.json_to_sheet(dependentes.length ? dependentes : []);

      XLSX.utils.book_append_sheet(wb, ws1, 'Principal');
      XLSX.utils.book_append_sheet(wb, ws2, 'Dependentes');

      // Salvar arquivo
      XLSX.writeFile(wb, 'cadastro_plano.xlsx');

      // Após salvar, manter os dados (não limpar automaticamente)
    }

    // Função para limpar formulário e tabela de dependentes
    function limparFormulario() {
      document.getElementById('formPrincipal').reset();
      const tabela = document.getElementById('tabelaDep');
      while (tabela.rows.length > 1) tabela.deleteRow(1);
    }

    // Pequena função de escape para prevenir quebra de HTML ao inserir valor no input readonly
    function escapeHtml(text) {
      if (!text) return '';
      return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    // Atualizar automaticamente o nome do titular nos dependentes
    document.getElementById('nome').addEventListener('input', function() {
      const novoNome = escapeHtml(this.value);
      document.querySelectorAll('#tabelaDep tr td:nth-child(2) input').forEach(input => {
        input.value = novoNome;
      });
    });

    // Associar eventos aos botões
    document.getElementById('btnAddDep').addEventListener('click', adicionarLinha);
    document.getElementById('btnGerar').addEventListener('click', gerarExcel);
    document.getElementById('btnLimpar').addEventListener('click', function() {
      if (confirm('Tem certeza que deseja limpar o formulário?')) limparFormulario();
    });