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

document.getElementById("btnExcel").addEventListener("click", gerarExcel);

function gerarExcel() {

    // --- PEGAR DADOS DO TITULAR ---
    const titular = {
        nome: document.getElementById("nomeTitular").value,
        cpf: document.getElementById("cpfTitular").value,
        nascimento: document.getElementById("nascimentoTitular").value,
        telefone: document.getElementById("telefoneTitular").value,
        email: document.getElementById("emailTitular").value,
        endereco: document.getElementById("enderecoTitular").value
    };

    // --- PEGAR DADOS DOS DEPENDENTES ---
    const dependentes = [];
    const linhas = document.querySelectorAll(".linha-dependente");

    linhas.forEach(linha => {
        dependentes.push({
            titular: titular.nome, // VINCULA AUTOMATICAMENTE
            nome: linha.querySelector(".dep-nome").value,
            nascimento: linha.querySelector(".dep-nascimento").value,
            cpf: linha.querySelector(".dep-cpf").value,
            parentesco: linha.querySelector(".dep-parentesco").value
        });
    });

    // --- MONTAR O EXCEL ---
    const wb = XLSX.utils.book_new();

    // Aba: Titular
    const wsTitular = XLSX.utils.json_to_sheet([titular]);
    XLSX.utils.book_append_sheet(wb, wsTitular, "Titular");

    // Aba: Dependentes
    const wsDependentes = XLSX.utils.json_to_sheet(dependentes);
    XLSX.utils.book_append_sheet(wb, wsDependentes, "Dependentes");

    // --- GERAR DOWNLOAD ---
    XLSX.writeFile(wb, "dados.xlsx");
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
